// Test the whole pipeline in Node:  node tests/run.mjs path/to/deck.pptx [--first] [--theme=svns]
// Prints each slide's detected type and options, then writes tests/out/<name>-refreshed.pptx
// containing every option for every visible slide (or just the first option with --first).

import JSZip from 'jszip';
import PptxGenJS from 'pptxgenjs';
import sharp from 'sharp';
import { readFileSync, mkdirSync, writeFileSync } from 'fs';
import { basename, dirname, join } from 'path';
import { fileURLToPath } from 'url';
import { parsePptx } from '../js/parse.js';
import { detect, extract, TYPES } from '../js/classify.js';
import { optionsFor, renderWith } from '../js/templates.js';
import { buildPptx, cropBox, finishFile, looksLikeGraphic, gradientSvg } from '../js/build.js';
import { STYLE, THEMES, DEFAULT_THEME } from '../js/style.js';

const here = dirname(fileURLToPath(import.meta.url));
const file = process.argv[2] || join(here, 'sample.pptx');
const firstOnly = process.argv.includes('--first');
const only = ((process.argv.find((a) => a.startsWith('--slides=')) || '').split('=')[1] || '').split(',').filter(Boolean).map(Number);
const themeKey = (process.argv.find((a) => a.startsWith('--theme=')) || '').split('=')[1] || DEFAULT_THEME;
const theme = THEMES[themeKey];
if (!theme) throw new Error(`Unknown theme "${themeKey}". Try: ${Object.keys(THEMES).join(', ')}`);
async function getGradient(spec) {
  const scale = Math.min(150, 1800 / Math.max(spec.w, spec.h));
  const w = Math.max(2, Math.round(spec.w * scale)), h = Math.max(2, Math.round(spec.h * scale));
  const buf = await sharp(Buffer.from(gradientSvg(spec, w, h))).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

const deck = await parsePptx(readFileSync(file), JSZip);

async function getImageData(src, boxW, boxH, fit) {
  const bytes = await deck.getImage(src.image);
  if (!bytes) return null;
  const img = sharp(Buffer.from(bytes)).rotate();
  const meta = await img.metadata();
  const r = cropBox(meta.width, meta.height, src.crop, boxW, boxH, fit);
  let pipe = sharp(Buffer.from(bytes)).rotate().extract({ left: r.x, top: r.y, width: Math.min(r.w, meta.width - r.x), height: Math.min(r.h, meta.height - r.y) });
  if (r.w > STYLE.maxImagePx) pipe = pipe.resize({ width: STYLE.maxImagePx });
  const buf = await pipe.flatten({ background: '#ffffff' }).jpeg({ quality: 86 }).toBuffer();
  let out = { data: 'image/jpeg;base64,' + buf.toString('base64') };
  if (fit === 'contain') {
    const ar = r.w / r.h;
    const w = Math.min(boxW, boxH * ar), h = w / ar;
    out = { ...out, w, h, dx: (boxW - w) / 2, dy: (boxH - h) / 2 };
  }
  return out;
}

// mark screenshots/logos the same way the page does
const seen = new Map();
for (const s of deck.slides) for (const e of s.elements) {
  if (e.kind !== 'pic') continue;
  if (!seen.has(e.image)) {
    const bytes = await deck.getImage(e.image);
    let g = false;
    try {
      const { data } = await sharp(Buffer.from(bytes)).ensureAlpha().resize(16, 16, { fit: 'fill' }).raw().toBuffer({ resolveWithObject: true });
      const px = (x, y) => Array.from(data.slice((y * 16 + x) * 4, (y * 16 + x) * 4 + 4));
      g = looksLikeGraphic([px(0, 0), px(15, 0), px(0, 15), px(15, 15)]);
    } catch { g = false; }
    seen.set(e.image, g);
  }
  e.graphic = seen.get(e.image);
}

const entries = [];
let section = 0;
for (const s of deck.slides) {
  if (s.hidden) continue;
  if (only.length && !only.includes(s.number)) { if (detect(s) === 'divider') section++; continue; }
  const type = detect(s);
  const model = extract(s, type);
  const opts = optionsFor(model);
  if (type === 'divider') section++;
  console.log(String(s.number).padStart(2), TYPES[type].padEnd(18), opts.map((o) => o.id).join(', ').padEnd(40), (model.title || model.subtitle || '').slice(0, 60));
  const ctx = { section, label: `${STYLE.report}  |  Report`, theme };
  for (const o of firstOnly ? opts.slice(0, 1) : opts) {
    const boxes = renderWith(o, model, ctx);
    boxes.push({ t: 'text', x: 0.43, y: 7.1, w: 6, h: 0.25, paras: [{ runs: [{ text: `Slide ${s.number}: ${TYPES[type]}, option "${o.name}"`, size: 8, color: theme.muted }], align: 'left', after: 0 }] });
    entries.push({ boxes, background: theme.bg });
  }
}

const pres = await buildPptx(entries, { PptxGenJS, getImageData, getGradient });
mkdirSync(join(here, 'out'), { recursive: true });
const out = join(here, 'out', basename(file, '.pptx') + `-${themeKey}-refreshed.pptx`);
writeFileSync(out, await finishFile(await pres.write({ outputType: 'arraybuffer' }), JSZip));
console.log('\nWrote', out, `(${entries.length} slides)`);
