// Browser image handling: crop to the slide's own crop plus the new frame,
// shrink large photos, and save as JPEG (keeps files light and avoids
// "picture can't be displayed" when slides are pasted into other decks).

import { STYLE } from './style.js';
import { cropBox, looksLikeGraphic } from './build.js';

const bitmaps = new Map();

async function bitmapFor(deck, path) {
  if (!bitmaps.has(path)) {
    bitmaps.set(path, (async () => {
      const bytes = await deck.getImage(path);
      if (!bytes) return null;
      const ext = path.split('.').pop().toLowerCase();
      const type = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp' }[ext];
      if (!type) return null; // EMF/WMF/TIFF can't be drawn in the browser
      try { return await createImageBitmap(new Blob([bytes], { type })); } catch { return null; }
    })());
  }
  return bitmaps.get(path);
}

function draw(bmp, r, maxW) {
  const scale = Math.min(1, maxW / r.w);
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(r.w * scale));
  canvas.height = Math.max(1, Math.round(r.h * scale));
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(bmp, r.x, r.y, r.w, r.h, 0, 0, canvas.width, canvas.height);
  return canvas;
}

// For the .pptx: returns { data, w, h, dx, dy } in PptxGenJS format, or null.
export function imageMaker(deck) {
  return async (src, boxW, boxH, fit) => {
    const bmp = await bitmapFor(deck, src.image);
    if (!bmp) return null;
    const r = cropBox(bmp.width, bmp.height, src.crop, boxW, boxH, fit);
    const data = draw(bmp, r, STYLE.maxImagePx).toDataURL('image/jpeg', 0.86).replace(/^data:/, '');
    if (fit !== 'contain') return { data };
    const ar = r.w / r.h, w = Math.min(boxW, boxH * ar), h = w / ar;
    return { data, w, h, dx: (boxW - w) / 2, dy: (boxH - h) / 2 };
  };
}

// For previews: a small object URL of the image after the slide's own crop.
export async function thumbFor(deck, src) {
  const bmp = await bitmapFor(deck, src.image);
  if (!bmp) return null;
  const r = cropBox(bmp.width, bmp.height, src.crop, 1, 1, 'none');
  const canvas = draw(bmp, r, 480);
  return new Promise((res) => canvas.toBlob((b) => res(b ? URL.createObjectURL(b) : null), 'image/jpeg', 0.8));
}

// Marks each picture element with .graphic (screenshot/logo) by sampling its corners.
export async function markGraphics(deck) {
  const seen = new Map();
  for (const s of deck.slides) for (const e of s.elements) {
    if (e.kind !== 'pic') continue;
    if (!seen.has(e.image)) {
      seen.set(e.image, (async () => {
        const bmp = await bitmapFor(deck, e.image);
        if (!bmp) return false;
        const c = document.createElement('canvas'); c.width = 16; c.height = 16;
        const ctx = c.getContext('2d', { willReadFrequently: true });
        ctx.drawImage(bmp, 0, 0, 16, 16);
        const px = (x, y) => Array.from(ctx.getImageData(x, y, 1, 1).data);
        return looksLikeGraphic([px(0, 0), px(15, 0), px(0, 15), px(15, 15)]);
      })());
    }
    e.graphic = await seen.get(e.image);
  }
}
