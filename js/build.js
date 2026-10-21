// Turns the boxes from templates.js into a .pptx with PptxGenJS.
// Pass in the PptxGenJS constructor and an image function for the platform
// (browser canvas or Node sharp), so this file runs in both.

import { STYLE } from './style.js';

const F = STYLE.fonts;

// Pixel rectangle to keep from an image, after the slide's own crop and a
// cover/contain fit into a w x h box.
export function cropBox(imgW, imgH, crop, boxW, boxH, fit) {
  const c = crop || { l: 0, t: 0, r: 0, b: 0 };
  let x = imgW * c.l, y = imgH * c.t;
  let w = imgW * (1 - c.l - c.r), h = imgH * (1 - c.t - c.b);
  if (fit === 'cover') {
    const target = boxW / boxH;
    if (w / h > target) { const nw = h * target; x += (w - nw) / 2; w = nw; }
    else { const nh = w / target; y += (h - nh) / 2; h = nh; }
  }
  // crops can run past the image edge (negative values in PowerPoint), so keep inside the image
  if (!isFinite(x)) x = 0;
  if (!isFinite(y)) y = 0;
  x = Math.min(Math.max(0, x), imgW - 1); y = Math.min(Math.max(0, y), imgH - 1);
  w = Math.min(isFinite(w) ? w : imgW, imgW - x); h = Math.min(isFinite(h) ? h : imgH, imgH - y);
  return { x: Math.round(x), y: Math.round(y), w: Math.max(1, Math.floor(w)), h: Math.max(1, Math.floor(h)) };
}

function runOpts(r, p, isLastInPara, isFirstInPara) {
  const o = { fontFace: F.body, fontSize: r.size, bold: !!r.bold, color: r.color, breakLine: isLastInPara };
  if (r.url) o.hyperlink = { url: r.url };
  if (isFirstInPara) {
    o.align = p.align || 'left';
    if (p.after) o.paraSpaceAfter = p.after;
    if (p.bullet) o.bullet = { characterCode: '25C6', indent: 14 };
  }
  return o;
}

function addText(slide, b) {
  const runs = [];
  b.paras.forEach((p, pi) => {
    p.runs.forEach((r, ri) => {
      runs.push({ text: r.text, options: runOpts(r, p, ri === p.runs.length - 1 && pi < b.paras.length - 1, ri === 0) });
    });
  });
  if (!runs.length) return;
  const o = { x: b.x, y: b.y, w: b.w, h: b.h, valign: b.valign || 'top', margin: Math.round((b.inset || 0) * 72), fit: 'none', isTextBox: true, fontFace: F.body };
  if (b.fill) o.fill = { color: b.fill };
  slide.addText(runs, o);
}

function axisText(b, o) {
  return {
    catAxisLabelColor: b.textColor, catAxisLabelFontFace: F.body, catAxisLabelFontSize: 9,
    valAxisLabelColor: b.textColor, valAxisLabelFontFace: F.body, valAxisLabelFontSize: 9,
    valGridLine: { color: b.gridColor, size: 0.5 }, catGridLine: { style: 'none' },
    catAxisLineShow: false, valAxisLineShow: false,
    legendFontFace: F.body, legendFontSize: 9, legendColor: b.legendColor,
    titleFontFace: F.body, titleFontSize: 12, titleColor: b.labelColor,
    ...o,
  };
}

function addChart(pres, slide, b) {
  const box = { x: b.x, y: b.y, w: b.w, h: b.h };
  const title = b.title ? { showTitle: true, title: b.title } : { showTitle: false };
  if (b.kind === 'line') {
    const data = b.series.map((s) => ({ name: s.name, labels: b.cats, values: s.values.map((v) => (v == null ? null : v)) }));
    slide.addChart(pres.charts.LINE, data, axisText(b, {
      ...box, ...title, chartColors: b.colors, lineSize: 2, lineDataSymbol: 'circle', lineDataSymbolSize: 5,
      showLegend: true, legendPos: 'b', valAxisLabelFormatCode: b.numFmt,
    }));
  } else if (b.kind === 'bar') {
    const s = b.series[0];
    slide.addChart(pres.charts.BAR, [{ name: s.name || 'Value', labels: b.cats, values: s.values.map((v) => v ?? 0) }], axisText(b, {
      ...box, ...title, barDir: b.horizontal ? 'bar' : 'col', chartColors: b.pointColors, barGapWidthPct: 60, ...(b.horizontal ? { catAxisOrientation: 'maxMin' } : {}),
      showLegend: false, showValue: true, dataLabelFormatCode: b.numFmt, dataLabelFontFace: F.body, dataLabelFontSize: 9, dataLabelColor: b.labelColor,
      dataLabelPosition: 'outEnd', valAxisLabelFormatCode: b.numFmt,
    }));
  } else if (b.kind === 'bar100') {
    const data = b.series.map((s) => ({ name: s.name, labels: b.cats, values: s.values.map((v) => v ?? 0) }));
    slide.addChart(pres.charts.BAR, data, axisText(b, {
      ...box, showTitle: false, barDir: 'bar', barGrouping: 'percentStacked', chartColors: b.colors, barGapWidthPct: 70,
      catAxisOrientation: 'maxMin', valAxisHidden: true, valGridLine: { style: 'none' }, catAxisLabelFontSize: 12, catAxisLabelColor: b.labelColor,
      showLegend: true, legendPos: 't', showValue: true, dataLabelFormatCode: b.numFmt, dataLabelPosition: 'ctr',
      dataLabelFontFace: F.body, dataLabelFontSize: 10, dataLabelColor: 'FFFFFF', dataLabelFontBold: true,
    }));
  } else if (b.kind === 'donut') {
    const s = b.series[0];
    slide.addChart(pres.charts.DOUGHNUT, [{ name: s.name || 'Value', labels: b.cats, values: s.values.map((v) => v ?? 0) }], axisText(b, {
      ...box, ...title, chartColors: b.pointColors, holeSize: 62, showLegend: true, legendPos: 'b',
      showPercent: true, showValue: false, dataLabelColor: 'FFFFFF', dataLabelFontFace: F.body, dataLabelFontSize: 9,
    }));
  }
}

function addTable(slide, b) {
  const rows = b.rows.map((r, i) => r.map((cell) => ({
    text: cell,
    options: i === 0
      ? { bold: true, color: b.onHead, fill: { color: b.head }, fontFace: F.body, fontSize: 9 }
      : { color: b.ink, fill: { color: i % 2 ? b.rowA : b.rowB }, fontFace: F.body, fontSize: 9 },
  })));
  slide.addTable(rows, { x: b.x, y: b.y, w: b.w, border: { type: 'solid', pt: 0.5, color: b.rowB }, valign: 'middle', margin: 3 });
}

// Gradient line for a box: angle 0 runs left to right, 90 top to bottom, 135 top-right to bottom-left.
export function gradientLine(angle, w, h) {
  const a = (angle * Math.PI) / 180, dx = Math.cos(a), dy = Math.sin(a);
  const half = (Math.abs(w * dx) + Math.abs(h * dy)) / 2;
  return { x1: w / 2 - dx * half, y1: h / 2 - dy * half, x2: w / 2 + dx * half, y2: h / 2 + dy * half };
}
// SVG text of a gradient box (used by Node via sharp, and as a fallback)
export function gradientSvg(spec, wPx, hPx) {
  const g = gradientLine(spec.angle, wPx, hPx);
  const stops = spec.stops.map(([c, a, p]) => `<stop offset="${p}" stop-color="#${c}" stop-opacity="${a}"/>`).join('');
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${wPx}" height="${hPx}"><defs><linearGradient id="g" gradientUnits="userSpaceOnUse" x1="${g.x1}" y1="${g.y1}" x2="${g.x2}" y2="${g.y2}">${stops}</linearGradient></defs><rect width="${wPx}" height="${hPx}" fill="url(#g)"/></svg>`;
}

// PptxGenJS links charts with absolute paths ("/ppt/charts/chart1.xml"). PowerPoint
// copes, but Keynote and some importers don't, so rewrite every link as relative.
export async function finishFile(arrayBuffer, JSZip) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const relFiles = Object.keys(zip.files).filter((p) => p.endsWith('.rels'));
  for (const rp of relFiles) {
    const owner = rp.replace('_rels/', '').replace(/\.rels$/, ''); // e.g. ppt/slides/slide1.xml
    const ownerDir = owner.includes('/') ? owner.slice(0, owner.lastIndexOf('/') + 1) : '';
    let xml = await zip.file(rp).async('string');
    xml = xml.replace(/Target="\/([^"]+)"/g, (_, abs) => {
      const from = ownerDir.split('/').filter(Boolean), to = abs.split('/');
      let i = 0;
      while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
      return `Target="${'../'.repeat(from.length - i)}${to.slice(i).join('/')}"`;
    });
    zip.file(rp, xml);
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// True when an image's four corners are white or transparent: a screenshot,
// chart or logo rather than a photo. rgba: 4 corner pixels as [r,g,b,a] arrays.
export function looksLikeGraphic(corners) {
  return corners.every(([r, g, b, a]) => a < 20 || (r > 235 && g > 235 && b > 235));
}

// entries: [{ boxes }] in slide order. getImageData(src, w, h, fit) -> { data, w, h, dx, dy } or null
export async function buildPptx(entries, { PptxGenJS, getImageData, getGradient }) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';
  pres.theme = { headFontFace: F.body, bodyFontFace: F.body };
  pres.title = 'HSBC Slide Refresh';
  for (const entry of entries) {
    const slide = pres.addSlide();
    slide.background = { color: entry.background || 'FFFFFF' };
    for (const b of entry.boxes) {
      if (b.t === 'rect') slide.addShape(pres.shapes.RECTANGLE, { x: b.x, y: b.y, w: b.w, h: b.h, fill: { color: b.fill }, line: { type: 'none' } });
      else if (b.t === 'ellipse') slide.addShape(pres.shapes.OVAL, { x: b.x, y: b.y, w: b.w, h: b.h, fill: { color: b.fill }, line: { type: 'none' } });
      else if (b.t === 'gradient') {
        const data = getGradient ? await getGradient(b) : null;
        if (data) slide.addImage({ data, x: b.x, y: b.y, w: b.w, h: b.h });
      }
      else if (b.t === 'text') addText(slide, b);
      else if (b.t === 'image') {
        const img = await getImageData(b.src, b.w, b.h, b.fit);
        if (img) slide.addImage({ data: img.data, x: b.x + (img.dx || 0), y: b.y + (img.dy || 0), w: img.w ?? b.w, h: img.h ?? b.h });
      } else if (b.t === 'chart') addChart(pres, slide, b);
      else if (b.t === 'table') addTable(slide, b);
    }
  }
  return pres;
}
