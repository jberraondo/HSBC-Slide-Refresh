// Layout options for each slide type. Each option turns a slide model into a
// list of simple boxes (rect, text, image, chart) measured in inches. The same
// list draws the preview on the page and builds the .pptx, so they always match.

import { STYLE } from './style.js';

const C = STYLE.colors;
const M = STYLE.margin.x;
const SW = STYLE.slide.w, SH = STYLE.slide.h;
const CW = SW - 2 * M; // content width
const TOP = 1.5;       // content starts below the title area
const BOTTOM = 6.95;   // content ends above the footer

// ---------- small builders ----------
const run = (text, size, color = C.text, bold = false, extra = {}) => ({ text, size, color, bold, ...extra });
const para = (runs, o = {}) => ({ runs: Array.isArray(runs) ? runs : [runs], align: o.align || 'left', after: o.after ?? 0, bullet: o.bullet || null });
const rect = (x, y, w, h, fill) => ({ t: 'rect', x, y, w, h, fill });
const text = (x, y, w, h, paras, o = {}) => ({ t: 'text', x, y, w, h, paras, valign: o.valign || 'top', inset: o.inset ?? 0, fill: o.fill || null });
const image = (x, y, w, h, src, fit = 'cover') => (src ? { t: 'image', x, y, w, h, src, fit } : rect(x, y, w, h, C.card));
const chart = (x, y, w, h, spec) => ({ t: 'chart', x, y, w, h, ...spec });
const card = (x, y, w, h, fill = C.card) => [rect(x, y, w, h, fill), rect(x, y, w, 0.06, C.dark)];

export function fmt(v) {
  if (v == null || isNaN(v)) return '';
  const a = Math.abs(v);
  if (a >= 1e6) return `${(v / 1e6).toFixed(1)}M`;
  if (a >= 1e3) return `${(v / 1e3).toFixed(1)}K`;
  return Number.isInteger(v) ? v.toLocaleString('en-GB') : v.toFixed(1);
}

function header(m, w = CW) {
  const out = [];
  if (m.title) out.push(text(M, 0.42, w, 0.34, [para(run(m.title, 14, C.text))]));
  if (m.subtitle) out.push(text(M, 0.74, w, 0.7, [para(run(m.subtitle, 20, C.text))]));
  return out;
}

const compColor = (label, c) => {
  const tok = c.split(' ')[0];
  if (/^(=|\+\/-)/.test(tok)) return C.grey;
  if (/cpv|cost/i.test(label)) return tok.startsWith('+') ? C.bad : C.good;
  return tok.startsWith('+') ? C.good : tok.startsWith('-') ? C.bad : C.grey;
};
const compPara = (c, size, color, signColor) => {
  const sp = c.indexOf(' ');
  const tok = sp > 0 ? c.slice(0, sp) : c, rest = sp > 0 ? c.slice(sp) : '';
  if (/^[+\-=]/.test(tok)) return para([run(tok, size, signColor, true), run(rest, size, color)], { after: 2 });
  return para(run(c, size, color), { after: 2 });
};

function factsAndBullets(m, x, y, w, h) {
  const out = [];
  let yy = y;
  if (m.facts && m.facts.length) {
    const half = Math.ceil(m.facts.length / 2);
    const cols = m.facts.length > 3 ? [m.facts.slice(0, half), m.facts.slice(half)] : [m.facts];
    cols.forEach((col, i) => out.push(text(x + i * (w / cols.length), yy, w / cols.length - 0.1, 0.8,
      col.map((t, j) => para(run(t, 10, j === 0 ? C.text : C.grey, j === 0), { after: 1 })))));
    yy += 0.85;
    out.push(rect(x, yy, w, 0.012, C.light));
    yy += 0.18;
  }
  if (m.bullets && m.bullets.length) {
    const size = m.bullets.join(' ').length > 900 ? 10 : 11;
    out.push(text(x, yy, w, y + h - yy, m.bullets.map((b) => (/:$/.test(b) && b.length < 120
      ? para(run(b.replace(/:$/, ''), size + 1, C.red, true), { after: 4 })
      : para(run(b, size, C.text), { after: 8, bullet: C.red })))));
  }
  return out;
}

// ================= dividers =================
const dividers = [
  {
    id: 'panel', name: 'Red panel with number',
    render(m, ctx) {
      const PW = 5.0;
      return [
        rect(0, 0, PW, SH, C.red),
        image(PW, 0, SW - PW, SH, m.image),
        rect(PW, 0, SW - PW, 0.08, C.dark),
        text(0.55, 2.35, PW - 0.9, 1.3, [para(run(String(ctx.section).padStart(2, '0'), 66, C.white, true))]),
        rect(0.55, 3.75, 0.9, 0.06, C.white),
        text(0.55, 3.95, PW - 0.9, 1.6, [para(run(m.title, 32, C.white))]),
        text(0.55, 6.75, PW - 0.9, 0.35, [para(run(ctx.label, 10, C.panelLabel))]),
      ];
    },
  },
  {
    id: 'band', name: 'Full photo with red band',
    render(m, ctx) {
      return [
        image(0, 0, SW, SH, m.image),
        rect(0, 4.75, 7.2, 1.75, C.red),
        text(0.55, 4.95, 1.4, 1.3, [para(run(String(ctx.section).padStart(2, '0'), 44, C.white, true))]),
        text(1.9, 4.95, 5.1, 1.0, [para(run(m.title, 30, C.white))], { valign: 'middle' }),
        text(1.9, 5.95, 5.1, 0.4, [para(run(ctx.label, 10, C.panelLabel))]),
      ];
    },
  },
  {
    id: 'clean', name: 'Clean title with photo strip',
    render(m, ctx) {
      return [
        image(SW - 4.2, 0, 4.2, SH, m.image),
        rect(SW - 4.2, 0, 0.08, SH, C.red),
        text(M + 0.3, 2.2, 7.5, 1.2, [para(run(String(ctx.section).padStart(2, '0'), 60, C.red, true))]),
        rect(M + 0.3, 3.55, 0.9, 0.06, C.red),
        text(M + 0.3, 3.75, 7.8, 1.6, [para(run(m.title, 36, C.text))]),
        text(M + 0.3, 6.75, 6, 0.35, [para(run(ctx.label, 10, C.grey))]),
      ];
    },
  },
];

// ================= contents =================
const contents = [
  {
    id: 'panel', name: 'Red panel list with photo',
    render(m, ctx) {
      const PW = 5.0, items = m.items.slice(0, 9);
      const out = [
        rect(0, 0, PW, SH, C.red), image(PW, 0, SW - PW, SH, m.image), rect(PW, 0, SW - PW, 0.08, C.dark),
        text(0.55, 0.85, PW - 0.9, 0.8, [para(run('Contents', 32, C.white))]), rect(0.55, 1.75, 0.9, 0.06, C.white),
      ];
      const step = Math.min(0.6, 4.4 / Math.max(items.length, 1));
      items.forEach((t, i) => {
        const y = 2.1 + i * step;
        out.push(text(0.55, y, PW - 0.9, step - 0.08, [para([run(String(i + 1).padStart(2, '0'), 16, C.white, true), run('   ' + t, 16, C.white)])]));
        if (i < items.length - 1) out.push(rect(0.55, y + step - 0.08, PW - 0.9, 0.01, C.panelRule));
      });
      out.push(text(0.55, 6.75, PW - 0.9, 0.35, [para(run(ctx.label, 10, C.panelLabel))]));
      return out;
    },
  },
  {
    id: 'list', name: 'Numbered list',
    render(m, ctx) {
      const items = m.items.slice(0, 10);
      const out = [text(M, 0.6, 6, 0.8, [para(run('Contents', 32, C.text))]), rect(M, 1.5, 0.9, 0.06, C.red)];
      const col = items.length > 5 ? 2 : 1, per = Math.ceil(items.length / col);
      items.forEach((t, i) => {
        const c = Math.floor(i / per), r = i % per;
        const x = M + c * (CW / 2), y = 2.0 + r * 0.85;
        out.push(text(x, y, CW / 2 - 0.4, 0.7, [para([run(String(i + 1).padStart(2, '0'), 24, C.red, true), run('   ' + t, 18, C.text)])], { valign: 'middle' }));
        out.push(rect(x, y + 0.72, CW / 2 - 0.4, 0.01, C.light));
      });
      return out;
    },
  },
];

// ================= headline figures =================
function statCard(f, x, y, w, h, o = {}) {
  const fg = o.fg || C.white;
  return [
    rect(x, y, w, h, o.fill || C.card), rect(x, y, w, 0.06, C.dark),
    text(x, y + 0.08, w, h - 0.08, [
      para(run(f.num, o.big || 30, fg, true)),
      para(run(f.label, 12, fg), { after: 8 }),
      ...f.comps.map((c) => compPara(c, 10, o.compColor || fg, o.signColor ? o.signColor(f.label, c) : fg)),
    ], { fill: o.inner === undefined ? C.red : o.inner, inset: 0.18 }),
  ];
}
const stats = [
  {
    id: 'cards', name: 'Red stat cards',
    render(m) {
      const out = header(m);
      const figs = m.figures.slice(0, 6);
      const hasCopy = (m.facts && m.facts.length) || (m.bullets && m.bullets.length);
      const gx = hasCopy ? 5.33 : M, gw = hasCopy ? SW - M - 5.33 : CW;
      const cols = figs.length === 4 || figs.length === 2 ? 2 : Math.min(3, figs.length), rows = Math.ceil(figs.length / cols);
      const gap = 0.16, w = (gw - gap * (cols - 1)) / cols, h = Math.min(2.5, (BOTTOM - TOP - gap * (rows - 1)) / rows);
      figs.forEach((f, i) => out.push(...statCard(f, gx + (i % cols) * (w + gap), TOP + Math.floor(i / cols) * (h + 0.2), w, h)));
      if (hasCopy) out.push(...factsAndBullets(m, M, TOP, 4.55, BOTTOM - TOP));
      return out;
    },
  },
  {
    id: 'row', name: 'Clean row of figures',
    render(m) {
      const out = header(m);
      const figs = m.figures.slice(0, 6), n = Math.max(figs.length, 1), cw = CW / n;
      out.push(rect(M, 1.45, 0.7, 0.05, C.red));
      figs.forEach((f, i) => out.push(text(M + i * cw, 1.8, cw - 0.12, 1.8, [
        para(run(f.num, n > 4 ? 26 : 30, C.text, true)), para(run(f.label, 11, C.grey), { after: 8 }),
        ...f.comps.map((c) => compPara(c, 9.5, C.grey, compColor(f.label, c))),
      ])));
      out.push(rect(M, 3.75, CW, 0.012, C.light));
      const hasFacts = m.facts && m.facts.length;
      if (hasFacts) out.push(text(M, 4.0, 2.2, 2.8, m.facts.map((t, j) => para(run(t, 10, j % 3 === 0 ? C.text : C.grey, j % 3 === 0), { after: 1 }))));
      const bx = hasFacts ? M + 2.4 : M, bw = SW - M - bx, bl = (m.bullets || []).slice(0, 3), cols = Math.max(bl.length, 1);
      bl.forEach((b, i) => out.push(text(bx + i * (bw / cols), 4.0, bw / cols - 0.25, 2.9, [para(run(b, 11, C.text), { bullet: C.red })])));
      return out;
    },
  },
  {
    id: 'block', name: 'Colour block',
    render(m) {
      const out = [rect(7.55, 0, SW - 7.55, SH, C.red), ...header(m, 6.9)];
      const figs = m.figures.slice(0, 6);
      const rows = Math.ceil(figs.length / 2), h = Math.min(2.1, 6.4 / Math.max(rows, 1));
      figs.forEach((f, i) => out.push(text(7.95 + (i % 2) * 2.7, 0.8 + Math.floor(i / 2) * h, 2.55, h - 0.1, [
        para(run(f.num, 30, C.white, true)), para(run(f.label, 12, C.white), { after: 8 }),
        ...f.comps.map((c) => compPara(c, 10, 'FBE3E5', compColor(f.label, c) === C.good ? C.white : compColor(f.label, c) === C.grey ? 'F2F2F2' : 'FFC9CE')),
      ])));
      out.push(...factsAndBullets(m, M, 1.7, 6.7, BOTTOM - 1.7));
      return out;
    },
  },
];

// ================= content cards =================
function contentCard(it, x, y, w, h, o = {}) {
  const ih = o.imageH || 2.05;
  // screenshots and logos (white or transparent corners) are fitted whole on a white panel; photos fill the space
  const isGraphic = !!(it.pic && it.pic.graphic);
  const out = [...card(x, y, w, h)];
  if (isGraphic) out.push(rect(x + 0.12, y + 0.18, w - 0.24, ih - 0.12, C.white), image(x + 0.2, y + 0.24, w - 0.4, ih - 0.24, it.pic, 'contain'));
  else out.push(image(x, y + 0.06, w, ih, it.pic));
  const titleRuns = (it.titleRuns && it.titleRuns.length ? it.titleRuns : [{ text: it.titleText }])
    .map((r) => run(r.text, o.titleSize || 11, C.text, true, r.url ? { url: r.url } : {}));
  out.push(text(x, y + ih + 0.13, w, h - ih - (it.stats.length ? 1.45 : 0.3), [
    para(titleRuns, { after: 5 }),
    ...(o.joinDesc ? [para(run(it.desc.join(' '), o.descSize || 9.5, C.soft))] : it.desc.map((d) => para(run(d, o.descSize || 9.5, C.soft), { after: 4 }))),
  ], { inset: 0.16 }));
  if (it.stats.length) {
    const inRow = o.statsInRow || 3, rowsN = Math.ceil(it.stats.length / inRow);
    const sy = y + h - 0.2 - rowsN * 0.55, cw = (w - 0.32) / inRow;
    out.push(rect(x + 0.16, sy - 0.1, w - 0.32, 0.01, C.light));
    it.stats.forEach((s, j) => out.push(text(x + 0.16 + (j % inRow) * cw, sy + Math.floor(j / inRow) * 0.55, cw, 0.55, [
      para(run(s.value, o.statSize || 12.5, j === 0 ? C.red : C.text, true)), para(run(s.label, 8.5, C.grey)),
    ])));
  }
  return out;
}
const cards = [
  {
    id: 'grid', name: 'Card grid',
    render(m) {
      // 1 to 3 cards: three slots across (empty slots ready for more); 4: four across; 5 to 8: two rows
      const out = header(m), items = m.items.slice(0, 8), n = items.length;
      const cols = n <= 3 ? 3 : n === 4 ? 4 : Math.ceil(n / 2), rows = n <= 4 ? 1 : 2, gap = cols > 3 ? 0.2 : 0.27;
      const w = (CW - gap * (cols - 1)) / cols, h = rows === 1 ? 5.38 : (5.38 - 0.2) / 2;
      const small = rows === 2 || cols > 3;
      items.forEach((it, i) => out.push(...contentCard(it, M + (i % cols) * (w + gap), 1.62 + Math.floor(i / cols) * (h + 0.2), w, h,
        small ? { imageH: rows === 2 ? 1.05 : 1.8, titleSize: 10, descSize: 8.5, statSize: 11, statsInRow: Math.min(3, Math.max(it.stats.length, 1)) } : {})));
      return out;
    },
  },
  {
    id: 'wide', name: 'Wide hero cards',
    render(m) {
      const out = header(m), items = m.items.slice(0, 2);
      const w = items.length === 1 ? 7.0 : (CW - 0.22) / 2, x0 = items.length === 1 ? (SW - w) / 2 : M;
      items.forEach((it, i) => out.push(...contentCard(it, x0 + i * (w + 0.22), 1.55, w, 5.4,
        { imageH: 2.45, titleSize: 13, descSize: 10, joinDesc: true, statsInRow: Math.min(Math.max(it.stats.length, 1), 5),
          statSize: it.stats.some((st) => st.value.length > 7) ? 12.5 : 18 })));
      return out;
    },
  },
  {
    id: 'feature', name: 'Single feature',
    when: (m) => m.items.length === 1,
    render(m) {
      const it = m.items[0], out = header(m);
      out.push(...card(M, 1.55, CW, 5.4), image(M, 1.61, 6.3, 5.34, it.pic));
      const titleRuns = (it.titleRuns.length ? it.titleRuns : [{ text: it.titleText }]).map((r) => run(r.text, 16, C.text, true, r.url ? { url: r.url } : {}));
      out.push(text(M + 6.5, 1.85, CW - 6.7, 2.6, [para(titleRuns, { after: 8 }), ...it.desc.map((d) => para(run(d, 11, C.soft), { after: 6 }))]));
      const inRow = Math.min(3, Math.max(it.stats.length, 1)), cw = (CW - 6.7) / inRow;
      out.push(rect(M + 6.5, 4.85, CW - 6.7, 0.01, C.light));
      it.stats.forEach((s, j) => out.push(text(M + 6.5 + (j % inRow) * cw, 5.0 + Math.floor(j / inRow) * 0.85, cw, 0.8, [
        para(run(s.value, 22, j === 0 ? C.red : C.text, true)), para(run(s.label, 10, C.grey))])));
      return out;
    },
  },
];

// ================= charts =================
const numFmt = (vals) => {
  const max = Math.max(...vals.filter((v) => v != null).map(Math.abs), 0);
  return max >= 1e6 ? '#,##0.0,,"M"' : max >= 1e4 ? '#,##0,"K"' : max <= 1 ? '0.0%' : '#,##0';
};
const lastIdx = (vals) => { for (let i = vals.length - 1; i >= 0; i--) if (vals[i] != null) return i; return -1; };

function lineSpec(c, o = {}) {
  // latest year first so it gets the red
  const series = [...c.series].sort((a, b) => (parseFloat(b.name) || 0) - (parseFloat(a.name) || 0));
  return {
    kind: 'line', cats: c.cats, series: series.slice().reverse(),
    colors: series.slice().reverse().map((s, i, arr) => (i === arr.length - 1 ? C.red : C.mid)),
    title: o.title === false ? '' : c.title, legend: true,
    numFmt: numFmt(c.series.flatMap((s) => s.values)),
  };
}
function lineStat(c) {
  const series = [...c.series].sort((a, b) => (parseFloat(b.name) || 0) - (parseFloat(a.name) || 0));
  const cur = series[0], prev = series[1];
  const i = lastIdx(cur.values);
  const comps = [];
  if (prev && i >= 0 && prev.values[i]) {
    const pct = (cur.values[i] / prev.values[i] - 1) * 100;
    comps.push(`${pct >= 0 ? '+' : ''}${pct.toFixed(0)}% on the same point in ${prev.name}`);
  }
  const label = c.title.replace(/\s*\d{2}\s*vs\s*\d{2}\s*$/i, '').trim() || c.title;
  return { num: fmt(cur.values[i]), label: `${label} to ${c.cats[i] || 'date'}`.trim(), comps };
}
function barSpec(c, o = {}) {
  const vals = c.series[0] ? c.series[0].values : [];
  return {
    kind: 'bar', horizontal: o.horizontal ?? false, cats: c.cats, series: [c.series[0] || { name: '', values: [] }],
    pointColors: c.cats.map((_, i) => STYLE.chartSeries[i % STYLE.chartSeries.length]),
    title: c.title, legend: false, numFmt: numFmt(vals), dataLabels: true,
  };
}
function donutSpec(c) {
  const vals = c.series[0] ? c.series[0].values : [];
  const pillarColors = c.cats.map((k, i) => STYLE.pillars[k.toLowerCase().trim()] || STYLE.chartSeries[i % STYLE.chartSeries.length]);
  return { kind: 'donut', cats: c.cats, series: [c.series[0] || { name: '', values: [] }], pointColors: pillarColors, title: c.title, legend: true, dataLabels: true, vals };
}
const isRate = (c) => /rate|%|percent/i.test(c.title) || (c.series[0] && c.series[0].values.every((v) => v == null || v < 100) && c.series[0].values.some((v) => v % 1));

function chartsLeftPanel(m) {
  return (m.facts && m.facts.length) || (m.bullets && m.bullets.length) || (m.figures && m.figures.length);
}

const chartOptions = [
  {
    id: 'statcards', name: 'Stat cards above each chart', when: (m) => m.charts.every((c) => c.kind === 'line'),
    render(m) {
      const out = header(m), list = m.charts.slice(0, 3), n = list.length, w = (CW - 0.22 * (n - 1)) / n;
      list.forEach((c, i) => {
        const x = M + i * (w + 0.22), f = lineStat(c);
        out.push(...statCard(f, x, 1.45, w, 1.75));
        out.push(...card(x, 3.35, w, 3.45), chart(x, 3.45, w, 3.3, lineSpec(c, { title: false })));
      });
      return out;
    },
  },
  {
    id: 'grid', name: 'Charts on cards',
    render(m) {
      const left = chartsLeftPanel(m);
      const out = header(m), list = m.charts.slice(0, 4), n = list.length;
      const gx = left ? 4.95 : M, gw = left ? SW - M - 4.95 : CW;
      if (left) out.push(...factsAndBullets(m, M, TOP, 4.2, BOTTOM - TOP));
      const cols = n === 1 ? 1 : 2, rows = Math.ceil(n / cols), gap = 0.15;
      const w = (gw - gap * (cols - 1)) / cols, h = (BOTTOM - 1.45 - gap * (rows - 1)) / rows;
      list.forEach((c, i) => {
        let x = gx + (i % cols) * (w + gap), y = 1.45 + Math.floor(i / cols) * (h + gap);
        if (n === 3 && i === 2 && !left) x = gx + (gw - w) / 2;
        const spec = c.kind === 'line' ? lineSpec(c) : c.kind === 'donut' || c.kind === 'pie' || c.kind === 'doughnut' ? donutSpec(c) : barSpec(c);
        out.push(...card(x, y, w, h), chart(x, y + 0.08, w, h - 0.12, spec));
      });
      return out;
    },
  },
  {
    id: 'hero', name: 'Hero chart with supporting charts', when: (m) => m.charts.length >= 2 && m.charts.length <= 3,
    render(m) {
      const out = header(m), [first, ...rest] = m.charts;
      const spec = (c) => (c.kind === 'line' ? lineSpec(c) : /pie|doughnut|donut/.test(c.kind) ? donutSpec(c) : barSpec(c));
      out.push(...card(M, 1.45, 7.9, 5.35), chart(M, 1.6, 7.9, 5.15, spec(first)));
      const h = rest.length === 1 ? 5.35 : 2.6;
      rest.forEach((c, i) => out.push(...card(8.53, 1.45 + i * (h + 0.15), 4.38, h), chart(8.53, 1.55 + i * (h + 0.15), 4.38, h - 0.15, spec(c))));
      return out;
    },
  },
  {
    id: 'share', name: 'Share of total bars', when: (m) => m.charts.length >= 2 && m.charts.every((c) => c.kind === 'bar') && m.charts.some((c) => !isRate(c)),
    render(m) {
      const out = header(m);
      const left = chartsLeftPanel(m);
      const gx = left ? 4.95 : M, gw = left ? SW - M - 4.95 : CW;
      if (left) out.push(...factsAndBullets(m, M, TOP, 4.2, BOTTOM - TOP));
      const shares = m.charts.filter((c) => !isRate(c)), rates = m.charts.filter(isRate);
      const cats = shares[0].cats;
      // one row per chart, one stacked series per category
      out.push(chart(gx, 1.4, gw, Math.min(1.25 * shares.length + 0.7, 4.0), {
        kind: 'bar100', horizontal: true,
        cats: shares.map((c) => c.title.replace(/\s*per\s+\w+\s*$/i, '')),
        series: cats.map((k, i) => ({ name: k, values: shares.map((c) => (c.series[0] ? c.series[0].values[i] : null)) })),
        colors: cats.map((_, i) => [C.red, C.mid, C.dark, C.light, C.gold][i % 5]),
        legend: true, dataLabels: true, numFmt: '[>=1000000]0.0,,"M";[>=1000]0.0,"K";0',
      }));
      rates.slice(0, 1).forEach((c) => {
        const y = 5.5;
        out.push(text(gx, y, gw, 0.35, [para(run(c.title.replace(/\s*per\s+\w+\s*$/i, ''), 13, C.text, true))]));
        const n = c.cats.length, w = (gw - 0.12 * (n - 1)) / n;
        c.cats.forEach((k, i) => out.push(text(gx + i * (w + 0.12), y + 0.42, w, 1.0, [
          para(run(`${(+c.series[0].values[i] || 0).toFixed(1)}%`, 26, C.white, true)), para(run(k, 11, C.white))],
        { fill: [C.red, C.mid, C.dark][i % 3], inset: 0.14 })));
      });
      return out;
    },
  },
  {
    id: 'donutcard', name: 'Donut on a card with copy', when: (m) => m.charts.length === 1 && /pie|doughnut|donut/.test(m.charts[0].kind),
    render(m) {
      const out = header(m);
      out.push(...factsAndBullets(m, M, 1.55, 6.9, BOTTOM - 1.55));
      out.push(...card(7.6, 1.5, 5.31, 5.35), chart(7.7, 1.65, 5.11, 5.1, donutSpec(m.charts[0])));
      return out;
    },
  },
];

// ================= learnings / text =================
const learnings = [
  {
    id: 'numbered', name: 'Numbered cards',
    render(m) {
      const out = header(m), items = m.items.slice(0, 4), w = m.image ? 7.05 : CW;
      if (m.image) out.push(image(7.73, 0, SW - 7.73, SH, m.image));
      const g = 0.12, h = (BOTTOM - 1.45 - g * (items.length - 1)) / items.length;
      items.forEach((it, i) => {
        const y = 1.45 + i * (h + g);
        out.push(...card(M, y, w, h), text(M, y + 0.06, 0.75, h - 0.06, [para(run(String(i + 1), 30, C.red, true), { align: 'center' })], { inset: 0.12 }));
        out.push(text(M + 0.7, y + 0.06, w - 0.75, h - 0.06, [
          ...(it.head ? [para(run(it.head, 12, C.red, true), { after: 5 })] : []), para(run(it.body, it.body.length > 450 ? 9 : 9.5, C.text)),
        ], { inset: 0.14 }));
      });
      return out;
    },
  },
  {
    id: 'columns', name: 'Side by side', when: (m) => m.items.length <= 3,
    render(m) {
      const out = header(m), items = m.items, n = items.length, w = (CW - 0.27 * (n - 1)) / n;
      items.forEach((it, i) => {
        const x = M + i * (w + 0.27);
        out.push(...card(x, 1.55, w, 5.4), text(x, 1.7, w, 0.9, [para(run(String(i + 1).padStart(2, '0'), 36, C.red, true))], { inset: 0.2 }));
        out.push(text(x, 2.6, w, 4.2, [...(it.head ? [para(run(it.head, 13, C.text, true), { after: 8 })] : []), para(run(it.body, 10.5, C.soft))], { inset: 0.2 }));
      });
      return out;
    },
  },
  {
    id: 'panel', name: 'Headings and bullets',
    render(m) {
      const out = header(m), ps = [];
      m.items.forEach((it) => { if (it.head) ps.push(para(run(it.head, 12.5, C.red, true), { after: 5 })); ps.push(para(run(it.body, 11, C.text), { after: 14, bullet: C.red })); });
      out.push(text(M, 1.55, m.image ? 7.0 : CW, 5.4, ps));
      if (m.image) out.push(image(7.73, 0, SW - 7.73, SH, m.image));
      return out;
    },
  },
];

const textOpts = [
  {
    id: 'bullets', name: 'Tidy bullets',
    render(m) {
      const out = header(m), w = m.image ? 7.0 : CW;
      out.push(...factsAndBullets(m, M, 1.55, w, BOTTOM - 1.55));
      if (m.image) out.push(...card(7.75, 1.5, 5.16, 5.4), image(7.75, 1.56, 5.16, 5.34, m.image));
      return out;
    },
  },
  {
    id: 'cards', name: 'Points as cards', when: (m) => m.bullets && m.bullets.length >= 2 && m.bullets.length <= 6,
    render(m) {
      const out = header(m), b = m.bullets, cols = b.length <= 3 ? b.length : Math.ceil(b.length / 2), rows = Math.ceil(b.length / cols);
      const w = (CW - 0.2 * (cols - 1)) / cols, h = (BOTTOM - 1.55 - 0.2 * (rows - 1)) / rows;
      b.forEach((t, i) => {
        const x = M + (i % cols) * (w + 0.2), y = 1.55 + Math.floor(i / cols) * (h + 0.2);
        out.push(...card(x, y, w, h), text(x, y + 0.1, w, h - 0.1, [para(run(t, 11, C.text))], { inset: 0.2 }));
      });
      return out;
    },
  },
];

const tableOpts = [
  {
    id: 'table', name: 'Styled tables with copy',
    render(m) {
      const out = header(m), left = (m.facts && m.facts.length) || (m.bullets && m.bullets.length);
      if (left) out.push(...factsAndBullets(m, M, TOP, 3.9, BOTTOM - TOP));
      const gx = left ? 4.6 : M, gw = SW - M - gx, n = Math.min(m.tables.length, 3), w = (gw - 0.2 * (n - 1)) / n;
      m.tables.slice(0, 3).forEach((rows, i) => out.push({ t: 'table', x: gx + i * (w + 0.2), y: TOP, w, h: BOTTOM - TOP, rows }));
      return out;
    },
  },
];

const cover = [
  {
    id: 'panel', name: 'Red panel cover',
    render(m, ctx) {
      const PW = m.image ? 6.2 : SW;
      return [
        rect(0, 0, PW, SH, C.red), ...(m.image ? [image(PW, 0, SW - PW, SH, m.image), rect(PW, 0, SW - PW, 0.08, C.dark)] : []),
        rect(0.7, 2.6, 0.9, 0.06, C.white),
        text(0.7, 2.8, Math.min(PW, 9) - 1.2, 3.6, [
          para(run(m.title, m.title.length > 30 ? 32 : 40, C.white), { after: 12 }),
          ...(m.subtitle ? [para(run(m.subtitle, 20, C.panelLabel))] : []),
        ]),
      ];
    },
  },
];

export const OPTIONS = {
  cover,
  divider: dividers, contents, stats, cards, charts: chartOptions,
  learnings, text: textOpts, table: tableOpts,
};

export function optionsFor(model) {
  return (OPTIONS[model.type] || textOpts).filter((o) => !o.when || o.when(model));
}
