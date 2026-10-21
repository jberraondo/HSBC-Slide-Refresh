// Test the whole pipeline in Node:  node tests/run.mjs path/to/deck.pptx [--first]
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
import { optionsFor } from '../js/templates.js';
import { buildPptx, cropBox, finishFile, looksLikeGraphic } from '../js/build.js';
import { STYLE } from '../js/style.js';

const here = dirname(fileURLToPath(import.meta.url));
const file = process.argv[2] || join(here, 'sample.pptx');
const firstOnly = process.argv.includes('--first');

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
  const type = detect(s);
  const model = extract(s, type);
  const opts = optionsFor(model);
  if (type === 'divider') section++;
  console.log(String(s.number).padStart(2), TYPES[type].padEnd(18), opts.map((o) => o.id).join(', ').padEnd(40), (model.title || model.subtitle || '').slice(0, 60));
  const ctx = { section, label: `${STYLE.report}  |  Report` };
  for (const o of firstOnly ? opts.slice(0, 1) : opts) {
    const boxes = o.render(model, ctx);
    boxes.push({ t: 'text', x: 0.43, y: 7.1, w: 6, h: 0.25, paras: [{ runs: [{ text: `Slide ${s.number}: ${TYPES[type]}, option "${o.name}"`, size: 8, color: '999999' }], align: 'left', after: 0 }] });
    entries.push({ boxes });
  }
}

const pres = await buildPptx(entries, { PptxGenJS, getImageData });
mkdirSync(join(here, 'out'), { recursive: true });
const out = join(here, 'out', basename(file, '.pptx') + '-refreshed.pptx');
writeFileSync(out, await finishFile(await pres.write({ outputType: 'arraybuffer' }), JSZip));
console.log('\nWrote', out, `(${entries.length} slides)`);
