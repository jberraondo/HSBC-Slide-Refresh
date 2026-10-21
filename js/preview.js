// Draws the boxes from templates.js as a small HTML preview of the slide.
// 1 inch = 100px on a 1333 x 750 canvas, scaled down with CSS.


const PX = 100;
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const pos = (b) => `left:${b.x * PX}px;top:${b.y * PX}px;width:${b.w * PX}px;height:${b.h * PX}px;`;

function textHtml(b) {
  const paras = b.paras.map((p) => {
    const runs = p.runs.map((r) => `<span style="font-size:${(r.size * PX) / 72}px;color:#${r.color};font-weight:${r.bold ? 700 : 300};${r.url ? 'text-decoration:underline;' : ''}">${esc(r.text)}</span>`).join('');
    const bullet = p.bullet ? `<span class="pv-bullet" style="color:#${p.bullet}">&#9670;</span>` : '';
    return `<p style="text-align:${p.align};margin:0 0 ${(p.after * PX) / 72}px ${p.bullet ? 22 : 0}px;position:relative">${bullet}${runs}</p>`;
  }).join('');
  const jc = b.valign === 'middle' ? 'center' : b.valign === 'bottom' ? 'flex-end' : 'flex-start';
  return `<div class="pv-text" style="${pos(b)}padding:${b.inset * PX}px;${b.fill ? `background:#${b.fill};` : ''}justify-content:${jc}">${paras}</div>`;
}

function chartSvg(b) {
  const w = b.w * PX, h = b.h * PX, pad = 30;
  const top = b.title ? 34 : 14, bottom = b.legend ? 46 : 26;
  let svg = '';
  if (b.title) svg += `<text x="${w / 2}" y="22" text-anchor="middle" font-size="15" fill="#${b.labelColor || "333333"}">${esc(b.title)}</text>`;
  if (b.kind === 'line') {
    const all = b.series.flatMap((s) => s.values).filter((v) => v != null);
    const max = Math.max(...all, 1), n = Math.max(b.cats.length, 2);
    const X = (i) => pad + (i * (w - 2 * pad)) / (n - 1), Y = (v) => h - bottom - ((h - top - bottom) * v) / max;
    for (let g = 0; g <= 4; g++) svg += `<line x1="${pad}" x2="${w - pad}" y1="${Y((max * g) / 4)}" y2="${Y((max * g) / 4)}" stroke="#${b.gridColor || "e3e3e3"}"/>`;
    b.series.forEach((s, si) => {
      const pts = s.values.map((v, i) => (v == null ? null : `${X(i)},${Y(v)}`)).filter(Boolean).join(' ');
      svg += `<polyline points="${pts}" fill="none" stroke="#${b.colors[si]}" stroke-width="3"/>`;
    });
  } else if (b.kind === 'bar') {
    const vals = b.series[0].values.map((v) => v || 0), max = Math.max(...vals, 1), n = vals.length;
    const bw = (w - 2 * pad) / n;
    vals.forEach((v, i) => {
      const bh = ((h - top - bottom) * v) / max;
      svg += `<rect x="${pad + i * bw + bw * 0.2}" y="${h - bottom - bh}" width="${bw * 0.6}" height="${bh}" fill="#${b.pointColors[i]}"/>`;
      svg += `<text x="${pad + i * bw + bw / 2}" y="${h - bottom + 16}" text-anchor="middle" font-size="11" fill="#${b.textColor || "767676"}">${esc(b.cats[i])}</text>`;
    });
  } else if (b.kind === 'bar100') {
    const rows = b.cats.length, rh = (h - 50) / Math.max(rows, 1), lx = 130;
    b.cats.forEach((cat, r) => {
      const tot = b.series.reduce((a, s) => a + (s.values[r] || 0), 0) || 1;
      let xx = lx;
      svg += `<text x="${lx - 10}" y="${40 + r * rh + rh / 2}" text-anchor="end" font-size="13" fill="#${b.labelColor || "333333"}">${esc(cat)}</text>`;
      b.series.forEach((s, si) => {
        const ww = ((w - lx - 10) * (s.values[r] || 0)) / tot;
        svg += `<rect x="${xx}" y="${30 + r * rh + rh * 0.15}" width="${Math.max(ww, 2)}" height="${rh * 0.6}" fill="#${b.colors[si]}"/>`;
        xx += ww;
      });
    });
  } else if (b.kind === 'donut') {
    const vals = b.series[0].values.map((v) => v || 0), tot = vals.reduce((a, v) => a + v, 0) || 1;
    const cx = w / 2, cy = top + (h - top - bottom) / 2, R = Math.min(w, h - top - bottom) / 2 - 6, r0 = R * 0.62;
    let a = -Math.PI / 2;
    vals.forEach((v, i) => {
      const a2 = a + (2 * Math.PI * v) / tot, large = a2 - a > Math.PI ? 1 : 0;
      const p = (ang, rad) => `${cx + rad * Math.cos(ang)},${cy + rad * Math.sin(ang)}`;
      svg += `<path d="M${p(a, R)} A${R},${R} 0 ${large} 1 ${p(a2, R)} L${p(a2, r0)} A${r0},${r0} 0 ${large} 0 ${p(a, r0)} Z" fill="#${b.pointColors[i]}"/>`;
      a = a2;
    });
  }
  if (b.legend) {
    const names = b.kind === 'donut' ? b.cats : b.series.map((s) => s.name);
    const cols = b.kind === 'donut' ? b.pointColors : b.colors;
    const y = b.kind === 'bar100' ? 14 : h - 14;
    const step = Math.min(160, (w - 20) / Math.max(names.length, 1));
    const start = w / 2 - (step * names.length) / 2;
    names.forEach((n, i) => {
      svg += `<rect x="${start + i * step}" y="${y - 9}" width="10" height="10" fill="#${cols[i]}"/><text x="${start + i * step + 14}" y="${y}" font-size="11" fill="#${b.legendColor || "333333"}">${esc(n)}</text>`;
    });
  }
  return `<svg class="pv-chart" style="${pos(b)}" viewBox="0 0 ${w} ${h}">${svg}</svg>`;
}

function tableHtml(b) {
  const rows = b.rows.map((r, i) => `<tr>${r.map((c) => `<td style="${i === 0 ? `background:#${b.head};color:#${b.onHead};font-weight:700;` : `background:#${i % 2 ? b.rowA : b.rowB};color:#${b.ink};`}">${esc(c)}</td>`).join('')}</tr>`).join('');
  return `<table class="pv-table" style="${pos(b)}height:auto">${rows}</table>`;
}

function rgba(hex, a) {
  const n = parseInt(hex, 16);
  return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`;
}
function cssGradient(b) {
  // CSS angles start at "to top" (0deg); ours start at "to the right"
  const stops = b.stops.map(([c, a, p]) => `${rgba(c, a)} ${Math.round(p * 100)}%`).join(',');
  return `linear-gradient(${b.angle + 90}deg,${stops})`;
}

// thumbs: Map from image path -> object URL (filled in lazily by the app)
export function renderPreview(boxes, thumbs) {
  const parts = boxes.map((b) => {
    if (b.t === 'rect') return `<div class="pv-box" style="${pos(b)}background:#${b.fill}"></div>`;
    if (b.t === 'ellipse') return `<div class="pv-box" style="${pos(b)}background:#${b.fill};border-radius:50%"></div>`;
    if (b.t === 'gradient') return `<div class="pv-box" style="${pos(b)}background:${cssGradient(b)}"></div>`;
    if (b.t === 'text') return textHtml(b);
    if (b.t === 'chart') return chartSvg(b);
    if (b.t === 'table') return tableHtml(b);
    if (b.t === 'image') {
      const url = thumbs.get(b.src.image);
      if (!url) return `<div class="pv-box pv-img-missing" data-img="${esc(b.src.image)}" style="${pos(b)}"></div>`;
      return `<div class="pv-box" style="${pos(b)}background:url('${url}') center/${b.fit === 'contain' ? 'contain' : 'cover'} no-repeat"></div>`;
    }
    return '';
  });
  return `<div class="pv-slide">${parts.join('')}</div>`;
}
