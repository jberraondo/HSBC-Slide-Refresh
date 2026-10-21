// Layout options for each slide type. Each option turns a slide model into a list
// of simple boxes (rect, gradient, text, image, chart, table) in inches, coloured
// from the chosen theme. The same list draws the preview and builds the .pptx.

import { STYLE } from './style.js';

const M = STYLE.margin.x;
const SW = STYLE.slide.w, SH = STYLE.slide.h;
const CW = SW - 2 * M; // content width
const TOP = 1.5;       // content starts below the title area
const BOTTOM = 6.95;   // content ends above the footer
const pad2 = (n) => String(n).padStart(2, '0');

// ---------- small builders ----------
const run = (text, size, color, bold = false, extra = {}) => ({ text, size, color, bold, ...extra });
const para = (runs, o = {}) => ({ runs: Array.isArray(runs) ? runs : [runs], align: o.align || 'left', after: o.after ?? 0, bullet: o.bullet || null });
const rect = (x, y, w, h, fill) => ({ t: 'rect', x, y, w, h, fill });
const ellipse = (x, y, w, h, fill) => ({ t: 'ellipse', x, y, w, h, fill });
const grad = (x, y, w, h, angle, stops) => ({ t: 'gradient', x, y, w, h, angle, stops });
const text = (x, y, w, h, paras, o = {}) => ({ t: 'text', x, y, w, h, paras, valign: o.valign || 'top', inset: o.inset ?? 0, fill: o.fill || null });
const chart = (x, y, w, h, spec) => ({ t: 'chart', x, y, w, h, ...spec });
const ca = (T) => T.cardAccent || T.accent; // highlight colour when sitting on a card
const card = (T, x, y, w, h) => [rect(x, y, w, h, T.surface), rect(x, y, w, 0.06, T.bar)];

function image(T, x, y, w, h, src, fit = 'cover') {
  return src ? { t: 'image', x, y, w, h, src, fit } : rect(x, y, w, h, T.surface);
}
// photo that fades into the background colour on one side ('left' = text sits on the left)
function photoFade(T, x, y, w, h, src, side = 'left', reach = 0.6) {
  if (!src) return [rect(x, y, w, h, T.surface)];
  const angle = side === 'left' ? 0 : side === 'right' ? 180 : side === 'bottom' ? 270 : 90;
  return [image(T, x, y, w, h, src), grad(x, y, w, h, angle, [[T.fade, 1, 0], [T.fade, 0.75, reach * 0.55], [T.fade, 0, reach]])];
}
// panel background: theme glow gradient if it has one, else a flat panel colour
function panelBg(T, x, y, w, h) {
  if (T.glow) return [rect(x, y, w, h, T.glow[1]), grad(x, y, w, h, 135, [[T.glow[0], 1, 0], [T.glow[1], 1, 0.75]])];
  return [rect(x, y, w, h, T.panel)];
}

export function fmt(v) {
  if (v == null || isNaN(v)) return '';
  const a = Math.abs(v);
  if (a >= 1e6) return `${(v / 1e6).toFixed(1)}M`;
  if (a >= 1e3) return `${(v / 1e3).toFixed(1)}K`;
  return Number.isInteger(v) ? v.toLocaleString('en-GB') : v.toFixed(1);
}

function header(m, T, w = CW) {
  const out = [];
  if (m.title) out.push(text(M, 0.42, w, 0.34, [para(run(m.title, 14, T.muted))]));
  if (m.subtitle) out.push(text(M, 0.74, w, 0.7, [para(run(m.subtitle, 20, T.text))]));
  return out;
}

const compColor = (T, label, c) => {
  const tok = c.split(' ')[0];
  if (/^(=|\+\/-)/.test(tok)) return T.muted;
  if (/cpv|cost/i.test(label)) return tok.startsWith('-') ? T.good : T.bad;
  return tok.startsWith('+') || /^\d/.test(tok) ? T.good : tok.startsWith('-') ? T.bad : T.muted;
};
const compPara = (c, size, color, signColor) => {
  const sp = c.indexOf(' ');
  const tok = sp > 0 ? c.slice(0, sp) : c, rest = sp > 0 ? c.slice(sp) : '';
  if (/^[+\-=]|^\d[\d.,]*%$/.test(tok)) return para([run(tok, size, signColor, true), run(rest, size, color)], { after: 2 });
  return para(run(c, size, color), { after: 2 });
};

// Largest font size (down to `min`) at which these paragraphs roughly fit a w x h box.
// Univers Light averages about 0.5em per character; lines are about 1.2em tall.
export function fitSize(texts, w, h, start, min = 7.5, gapPt = 8) {
  for (let size = start; size > min; size -= 0.5) {
    const perLine = Math.max(1, (w * 72) / (size * 0.56));
    const lines = texts.reduce((n, t) => n + Math.max(1, Math.ceil(t.length / perLine)), 0);
    if ((lines * size * 1.2 + texts.length * gapPt) / 72 <= h) return size;
  }
  return min;
}

function copyParas(m, T, size, color = T.text) {
  return (m.bullets || []).map((b) => (/:$/.test(b) && b.length < 120
    ? para(run(b.replace(/:$/, ''), size + 1, T.accent, true), { after: 4 })
    : para(run(b, size, color), { after: 8, bullet: T.accent })));
}

function factsAndBullets(m, T, x, y, w, h) {
  const out = [];
  let yy = y;
  if (m.facts && m.facts.length) {
    const half = Math.ceil(m.facts.length / 2);
    const cols = m.facts.length > 3 ? [m.facts.slice(0, half), m.facts.slice(half)] : [m.facts];
    cols.forEach((col, i) => out.push(text(x + i * (w / cols.length), yy, w / cols.length - 0.1, 0.8,
      col.map((t, j) => para(run(t, 10, j === 0 ? T.text : T.muted, j === 0), { after: 1 })))));
    yy += 0.85;
    out.push(rect(x, yy, w, 0.012, T.line));
    yy += 0.18;
  }
  if (m.bullets && m.bullets.length) {
    const size = fitSize(m.bullets, w - 0.25, y + h - yy, 11);
    out.push(text(x, yy, w, y + h - yy, copyParas(m, T, size)));
  }
  return out;
}

// ================= cover =================
const cover = [
  {
    id: 'panel', name: 'Panel and photo',
    render(m, { theme: T }) {
      const PW = m.image ? 6.2 : SW;
      return [
        ...panelBg(T, 0, 0, PW, SH), ...(m.image ? [image(T, PW, 0, SW - PW, SH, m.image), rect(PW, 0, SW - PW, 0.08, T.bar)] : []),
        rect(0.7, 2.6, 0.9, 0.06, T.onPanel),
        text(0.7, 2.8, Math.min(PW, 9) - 1.2, 3.6, [
          para(run(m.title, m.title.length > 30 ? 32 : 40, T.onPanel), { after: 12 }),
          ...(m.subtitle ? [para(run(m.subtitle, 20, T.panelMuted))] : []),
        ]),
      ];
    },
  },
  {
    id: 'fade', name: 'Full photo, fade to title', when: (m) => !!m.image,
    render(m, { theme: T }) {
      return [
        rect(0, 0, SW, SH, T.fade), ...photoFade(T, 0, 0, SW, SH, m.image, 'left', 0.75),
        text(0.7, 2.4, 7.2, 3.4, [para(run(m.title, 44, T.dark ? 'FFFFFF' : T.text), { after: 10 }), ...(m.subtitle ? [para(run(m.subtitle, 22, T.dark ? 'FFFFFF' : T.text))] : [])]),
        rect(0.7, 2.2, 0.9, 0.06, T.accent),
      ];
    },
  },
  {
    id: 'glow', name: 'Glow, no photo',
    render(m, { theme: T }) {
      return [
        ...panelBg(T, 0, 0, SW, SH),
        text(0.9, 2.5, 9, 3, [para(run(m.title, 48, T.onPanel), { after: 10 }), ...(m.subtitle ? [para(run(m.subtitle, 22, T.panelMuted))] : [])]),
        rect(0.9, 2.3, 1.0, 0.07, T.onPanel),
      ];
    },
  },
  {
    id: 'split', name: 'Title left, photo right', when: (m) => !!m.image,
    render(m, { theme: T }) {
      return [
        rect(0, 0, SW, SH, T.bg), image(T, SW / 2, 0, SW / 2, SH, m.image), rect(SW / 2 - 0.08, 0, 0.08, SH, T.accent),
        text(0.7, 2.4, SW / 2 - 1.3, 3.4, [para(run(m.title, 40, T.text), { after: 10 }), ...(m.subtitle ? [para(run(m.subtitle, 20, T.muted))] : [])]),
        rect(0.7, 2.2, 0.9, 0.06, T.accent),
      ];
    },
  },
];

// ================= dividers =================
const num = (ctx) => pad2(ctx.section);
const dividers = [
  {
    id: 'panel', name: 'Panel with number',
    render(m, ctx) {
      const T = ctx.theme, PW = m.image ? 5.0 : SW;
      return [
        ...panelBg(T, 0, 0, PW, SH), ...(m.image ? [image(T, PW, 0, SW - PW, SH, m.image), rect(PW, 0, SW - PW, 0.08, T.bar)] : []),
        text(0.55, 2.35, Math.min(PW, 6) - 0.9, 1.3, [para(run(num(ctx), 66, T.onPanel, true))]),
        rect(0.55, 3.75, 0.9, 0.06, T.onPanel),
        text(0.55, 3.95, PW - 0.9, 1.6, [para(run(m.title, 32, T.onPanel))]),
        text(0.55, 6.75, PW - 0.9, 0.35, [para(run(ctx.label, 10, T.panelMuted))]),
      ];
    },
  },
  {
    id: 'glow', name: 'Glow with big title',
    render(m, ctx) {
      const T = ctx.theme;
      return [
        ...panelBg(T, 0, 0, SW, SH),
        text(0.9, 2.6, 2.0, 1.1, [para(run(num(ctx), 28, T.panelMuted, true))]),
        text(0.9, 3.25, 11, 1.4, [para(run(m.title, 44, T.onPanel, true))]),
        text(0.9, 6.75, 8, 0.35, [para(run(ctx.label, 10, T.panelMuted))]),
      ];
    },
  },
  {
    id: 'fade', name: 'Photo fading into title', when: (m) => !!m.image,
    render(m, ctx) {
      const T = ctx.theme, ink = T.dark ? 'FFFFFF' : T.text;
      return [
        rect(0, 0, SW, SH, T.fade), ...photoFade(T, 3.2, 0, SW - 3.2, SH, m.image, 'left', 0.6),
        text(0.7, 2.5, 7, 1.0, [para(run(num(ctx), 40, T.accent, true))]),
        text(0.7, 3.45, 7.5, 1.4, [para(run(m.title, 38, ink))]),
        text(0.7, 6.75, 8, 0.35, [para(run(ctx.label, 10, T.muted))]),
      ];
    },
  },
  {
    id: 'band', name: 'Full photo with band', when: (m) => !!m.image,
    render(m, ctx) {
      const T = ctx.theme;
      return [
        image(T, 0, 0, SW, SH, m.image), rect(0, 4.75, 7.2, 1.75, T.accent),
        text(0.55, 4.95, 1.4, 1.3, [para(run(num(ctx), 44, T.onAccent, true))]),
        text(1.9, 4.95, 5.1, 1.0, [para(run(m.title, 30, T.onAccent))], { valign: 'middle' }),
        text(1.9, 5.95, 5.1, 0.4, [para(run(ctx.label, 10, T.onAccent))]),
      ];
    },
  },
  {
    id: 'clean', name: 'Clean title with photo strip', when: (m) => !!m.image,
    render(m, ctx) {
      const T = ctx.theme;
      return [
        rect(0, 0, SW, SH, T.bg), image(T, SW - 4.2, 0, 4.2, SH, m.image), rect(SW - 4.2, 0, 0.08, SH, T.accent),
        text(M + 0.3, 2.2, 7.5, 1.2, [para(run(num(ctx), 60, T.accent, true))]),
        rect(M + 0.3, 3.55, 0.9, 0.06, T.accent),
        text(M + 0.3, 3.75, 7.8, 1.6, [para(run(m.title, 36, T.text))]),
        text(M + 0.3, 6.75, 6, 0.35, [para(run(ctx.label, 10, T.muted))]),
      ];
    },
  },
];

// ================= contents =================
const contents = [
  {
    id: 'panel', name: 'Panel list with photo',
    render(m, ctx) {
      const T = ctx.theme, PW = 5.0, items = m.items.slice(0, 9);
      const out = [...panelBg(T, 0, 0, PW, SH), image(T, PW, 0, SW - PW, SH, m.image), rect(PW, 0, SW - PW, 0.08, T.bar),
        text(0.55, 0.85, PW - 0.9, 0.8, [para(run('Contents', 32, T.onPanel))]), rect(0.55, 1.75, 0.9, 0.06, T.onPanel)];
      const step = Math.min(0.6, 4.4 / Math.max(items.length, 1));
      items.forEach((t, i) => {
        const y = 2.1 + i * step;
        out.push(text(0.55, y, PW - 0.9, step - 0.08, [para([run(pad2(i + 1), 16, T.onPanel, true), run('   ' + t, 16, T.onPanel)])]));
        if (i < items.length - 1) out.push(rect(0.55, y + step - 0.08, PW - 0.9, 0.01, T.panelRule));
      });
      out.push(text(0.55, 6.75, PW - 0.9, 0.35, [para(run(ctx.label, 10, T.panelMuted))]));
      return out;
    },
  },
  {
    id: 'grid', name: 'Photo and numbered grid',
    render(m, ctx) {
      const T = ctx.theme, items = m.items.slice(0, 10);
      const out = [rect(0, 0, SW, SH, T.bg), text(M, 0.7, 3.6, 1.2, [para(run('What you will see today', 30, T.text))]), image(T, M, 2.1, 3.5, 4.6, m.image)];
      const col = 2, per = Math.ceil(items.length / col), x0 = 4.6, cw = (SW - M - x0) / col;
      items.forEach((t, i) => {
        const c = Math.floor(i / per), r = i % per, x = x0 + c * cw, y = 0.75 + r * (6.0 / Math.max(per, 1));
        out.push(text(x, y, cw - 0.4, 0.8, [para(run(pad2(i + 1), 22, T.text, true)), para(run(t, 12, T.muted))]));
        out.push(rect(x, y + 0.88, cw - 0.4, 0.015, T.accent));
      });
      return out;
    },
  },
  {
    id: 'list', name: 'Numbered list',
    render(m, ctx) {
      const T = ctx.theme, items = m.items.slice(0, 10);
      const out = [rect(0, 0, SW, SH, T.bg), text(M, 0.6, 6, 0.8, [para(run('Contents', 32, T.text))]), rect(M, 1.5, 0.9, 0.06, T.accent)];
      const col = items.length > 5 ? 2 : 1, per = Math.ceil(items.length / col);
      items.forEach((t, i) => {
        const c = Math.floor(i / per), r = i % per, x = M + c * (CW / 2), y = 2.0 + r * 0.85;
        out.push(text(x, y, CW / 2 - 0.4, 0.7, [para([run(pad2(i + 1), 24, T.accent, true), run('   ' + t, 18, T.text)])], { valign: 'middle' }));
        out.push(rect(x, y + 0.72, CW / 2 - 0.4, 0.01, T.line));
      });
      return out;
    },
  },
];

// ================= headline figures =================
function statCard(T, f, x, y, w, h, o = {}) {
  return [
    rect(x, y, w, h, T.surface), rect(x, y, w, 0.06, T.bar),
    text(x, y + 0.08, w, h - 0.08, [
      para(run(f.num, o.big || 30, T.statNum, true)),
      para(run(f.label, 12, T.onStat), { after: 8 }),
      ...f.comps.map((c) => compPara(c, 10, T.onStat, T.statFill === T.surface || T.statFill === 'FFFFFF' ? compColor(T, f.label, c) : T.onStat)),
    ], { fill: T.statFill, inset: 0.18 }),
  ];
}
const withCopy = (m) => (m.facts && m.facts.length) || (m.bullets && m.bullets.length);
const stats = [
  {
    id: 'cards', name: 'Stat cards',
    render(m, { theme: T }) {
      const out = header(m, T), figs = m.figures.slice(0, 6), hasCopy = withCopy(m);
      const gx = hasCopy ? 5.33 : M, gw = hasCopy ? SW - M - 5.33 : CW;
      const cols = figs.length === 4 || figs.length === 2 ? 2 : Math.min(3, figs.length), rows = Math.ceil(figs.length / cols);
      const gap = 0.16, w = (gw - gap * (cols - 1)) / cols, h = Math.min(2.5, (BOTTOM - TOP - gap * (rows - 1)) / rows);
      figs.forEach((f, i) => out.push(...statCard(T, f, gx + (i % cols) * (w + gap), TOP + Math.floor(i / cols) * (h + 0.2), w, h)));
      if (hasCopy) out.push(...factsAndBullets(m, T, M, TOP, 4.55, BOTTOM - TOP));
      return out;
    },
  },
  {
    id: 'hero', name: 'One big number',
    render(m, { theme: T }) {
      const out = header(m, T), [first, ...rest] = m.figures;
      out.push(...panelBg(T, M, TOP, 5.2, BOTTOM - TOP));
      out.push(text(M, TOP + 0.6, 5.2, 4, [para(run(first.num, 80, T.onPanel, true)), para(run(first.label, 20, T.onPanel), { after: 12 }),
        ...first.comps.map((c) => compPara(c, 13, T.panelMuted, T.onPanel))], { inset: 0.4 }));
      const x = M + 5.5, w = SW - M - x, rowsH = (BOTTOM - TOP) / Math.max(rest.length + (withCopy(m) ? 1 : 0), 1);
      rest.slice(0, 5).forEach((f, i) => {
        const y = TOP + i * rowsH;
        out.push(text(x, y, 2.6, rowsH - 0.1, [para(run(f.num, 28, T.text, true)), para(run(f.label, 11, T.muted))]));
        out.push(text(x + 2.7, y + 0.05, w - 2.7, rowsH - 0.1, f.comps.map((c) => compPara(c, 10, T.muted, compColor(T, f.label, c)))));
        out.push(rect(x, y + rowsH - 0.08, w, 0.01, T.line));
      });
      if (withCopy(m)) { const hh = BOTTOM - (TOP + rest.slice(0, 5).length * rowsH); out.push(text(x, BOTTOM - hh, w, hh, copyParas(m, T, fitSize(m.bullets || [], w - 0.25, hh, 10)))); }
      return out;
    },
  },
  {
    id: 'row', name: 'Clean row of figures',
    render(m, { theme: T }) {
      const out = header(m, T), figs = m.figures.slice(0, 6), n = Math.max(figs.length, 1), cw = CW / n;
      out.push(rect(M, 1.45, 0.7, 0.05, T.accent));
      figs.forEach((f, i) => out.push(text(M + i * cw, 1.8, cw - 0.12, 1.8, [
        para(run(f.num, n > 4 ? 26 : 30, T.text, true)), para(run(f.label, 11, T.muted), { after: 8 }),
        ...f.comps.map((c) => compPara(c, 9.5, T.muted, compColor(T, f.label, c))),
      ])));
      out.push(rect(M, 3.75, CW, 0.012, T.line));
      const hasFacts = m.facts && m.facts.length;
      if (hasFacts) out.push(text(M, 4.0, 2.2, 2.8, m.facts.map((t, j) => para(run(t, 10, j % 3 === 0 ? T.text : T.muted, j % 3 === 0), { after: 1 }))));
      const bx = hasFacts ? M + 2.4 : M, bw = SW - M - bx, bl = m.bullets || [], cols = Math.min(Math.max(bl.length, 1), 3);
      const per = Math.ceil(bl.length / cols);
      for (let c = 0; c < cols; c++) {
        const part = bl.slice(c * per, (c + 1) * per), cw = bw / cols - 0.25, size = fitSize(bl.slice(0, per), cw - 0.25, 2.9, 11);
        out.push(text(bx + c * (bw / cols), 4.0, cw, 2.9, part.map((b) => para(run(b, size, T.text), { after: 6, bullet: T.accent }))));
      }
      return out;
    },
  },
  {
    id: 'rows', name: 'Figure list',
    render(m, { theme: T }) {
      const out = header(m, T), figs = m.figures.slice(0, 6), hasCopy = withCopy(m);
      const x = hasCopy ? 5.33 : M, w = SW - M - x, rh = (BOTTOM - TOP) / Math.max(figs.length, 1);
      figs.forEach((f, i) => {
        const y = TOP + i * rh;
        out.push(rect(x, y, 0.08, rh - 0.12, T.accent));
        out.push(text(x + 0.25, y, 2.8, rh - 0.1, [para(run(f.num, Math.min(28, rh * 26), T.text, true)), para(run(f.label, 10, T.muted))], { valign: 'middle' }));
        out.push(text(x + 3.2, y, w - 3.2, rh - 0.1, f.comps.map((c) => compPara(c, 10, T.muted, compColor(T, f.label, c))), { valign: 'middle' }));
      });
      if (hasCopy) out.push(...factsAndBullets(m, T, M, TOP, 4.55, BOTTOM - TOP));
      return out;
    },
  },
  {
    id: 'block', name: 'Colour block',
    render(m, { theme: T }) {
      const out = [...panelBg(T, 7.55, 0, SW - 7.55, SH), ...header(m, T, 6.9)];
      const figs = m.figures.slice(0, 6), rows = Math.ceil(figs.length / 2), h = Math.min(2.1, 6.4 / Math.max(rows, 1));
      figs.forEach((f, i) => out.push(text(7.95 + (i % 2) * 2.7, 0.8 + Math.floor(i / 2) * h, 2.55, h - 0.1, [
        para(run(f.num, 30, T.onPanel, true)), para(run(f.label, 12, T.onPanel), { after: 8 }),
        ...f.comps.map((c) => compPara(c, 10, T.panelMuted, T.onPanel)),
      ])));
      out.push(...factsAndBullets(m, T, M, 1.7, 6.7, BOTTOM - 1.7));
      return out;
    },
  },
];

// ================= content cards =================
function titleRuns(it, size, color) {
  return (it.titleRuns && it.titleRuns.length ? it.titleRuns : [{ text: it.titleText }]).map((r) => run(r.text, size, color, true, r.url ? { url: r.url } : {}));
}
function contentCard(T, it, x, y, w, h, o = {}) {
  const ih = o.imageH || 2.05;
  const out = [...card(T, x, y, w, h)];
  if (it.pic && it.pic.graphic) out.push(rect(x + 0.12, y + 0.18, w - 0.24, ih - 0.12, 'FFFFFF'), image(T, x + 0.2, y + 0.24, w - 0.4, ih - 0.24, it.pic, 'contain'));
  else out.push(image(T, x, y + 0.06, w, ih, it.pic));
  out.push(text(x, y + ih + 0.13, w, h - ih - (it.stats.length ? 1.45 : 0.3), [
    para(titleRuns(it, o.titleSize || 11, T.onSurface), { after: 5 }),
    ...(o.joinDesc ? [para(run(it.desc.join(' '), o.descSize || 9.5, T.soft))] : it.desc.map((d) => para(run(d, o.descSize || 9.5, T.soft), { after: 4 }))),
  ], { inset: 0.16 }));
  if (it.stats.length) {
    const inRow = o.statsInRow || 3, rowsN = Math.ceil(it.stats.length / inRow);
    const sy = y + h - 0.2 - rowsN * 0.55, cw = (w - 0.32) / inRow;
    out.push(rect(x + 0.16, sy - 0.1, w - 0.32, 0.01, T.line));
    it.stats.forEach((s, j) => out.push(text(x + 0.16 + (j % inRow) * cw, sy + Math.floor(j / inRow) * 0.55, cw, 0.55, [
      para(run(s.value, o.statSize || 12.5, j === 0 ? ca(T) : T.onSurface, true)), para(run(s.label, 8.5, T.muted)),
    ])));
  }
  return out;
}
const cards = [
  {
    id: 'grid', name: 'Card grid',
    render(m, { theme: T }) {
      // 1 to 3 cards: three slots across (empty slots ready for more); 4: four across; 5 to 8: two rows
      const out = header(m, T), items = m.items.slice(0, 8), n = items.length;
      const cols = n <= 3 ? 3 : n === 4 ? 4 : Math.ceil(n / 2), rows = n <= 4 ? 1 : 2, gap = cols > 3 ? 0.2 : 0.27;
      const w = (CW - gap * (cols - 1)) / cols, h = rows === 1 ? 5.38 : (5.38 - 0.2) / 2, small = rows === 2 || cols > 3;
      items.forEach((it, i) => out.push(...contentCard(T, it, M + (i % cols) * (w + gap), 1.62 + Math.floor(i / cols) * (h + 0.2), w, h,
        small ? { imageH: rows === 2 ? 1.05 : 1.8, titleSize: 10, descSize: 8.5, statSize: 11, statsInRow: Math.min(3, Math.max(it.stats.length, 1)) } : {})));
      return out;
    },
  },
  {
    id: 'overlay', name: 'Photo cards with text on top', when: (m) => m.items.length <= 4,
    render(m, { theme: T }) {
      const out = header(m, T), items = m.items, n = items.length, gap = 0.2, w = (CW - gap * (n - 1)) / n, y = 1.55, h = 5.4;
      items.forEach((it, i) => {
        const x = M + i * (w + gap);
        out.push(image(T, x, y, w, h, it.pic), grad(x, y, w, h, 270, [['0B0B0B', 0.95, 0], ['0B0B0B', 0.7, 0.35], ['0B0B0B', 0, 0.7]]), rect(x, y, w, 0.06, T.accent));
        const statLine = it.stats.slice(0, 3).map((s) => para([run(s.value, 15, 'FFFFFF', true), run('  ' + s.label, 9, 'D9D9D9')], { after: 2 }));
        out.push(text(x, y + h * 0.42, w, h * 0.58, [para(titleRuns(it, 13, 'FFFFFF'), { after: 6 }), para(run(it.desc.join(' '), 9, 'E6E6E6'), { after: 8 }), ...statLine], { inset: 0.2, valign: 'bottom' }));
      });
      return out;
    },
  },
  {
    id: 'rows', name: 'Image and text rows', when: (m) => m.items.length >= 2 && m.items.length <= 4,
    render(m, { theme: T }) {
      const out = header(m, T), items = m.items, n = items.length, gap = 0.15, h = (BOTTOM - 1.5 - gap * (n - 1)) / n;
      items.forEach((it, i) => {
        const y = 1.5 + i * (h + gap), iw = h * 1.6;
        out.push(...card(T, M, y, CW, h), image(T, M, y + 0.06, iw, h - 0.06, it.pic, it.pic && it.pic.graphic ? 'contain' : 'cover'));
        out.push(text(M + iw + 0.1, y + 0.06, 5.2, h - 0.06, [para(titleRuns(it, 12, T.onSurface), { after: 4 }), para(run(it.desc.join(' '), 9.5, T.soft))], { inset: 0.12, valign: 'middle' }));
        const sx = M + iw + 5.4, sw = SW - M - sx, k = Math.max(Math.min(it.stats.length, 4), 1);
        it.stats.slice(0, 4).forEach((s, j) => out.push(text(sx + j * (sw / k), y + 0.06, sw / k, h - 0.06, [para(run(s.value, 15, j === 0 ? ca(T) : T.onSurface, true)), para(run(s.label, 8.5, T.muted))], { valign: 'middle' })));
      });
      return out;
    },
  },
  {
    id: 'wide', name: 'Wide hero cards',
    render(m, { theme: T }) {
      const out = header(m, T), items = m.items.slice(0, 2);
      const w = items.length === 1 ? 7.0 : (CW - 0.22) / 2, x0 = items.length === 1 ? (SW - w) / 2 : M;
      items.forEach((it, i) => out.push(...contentCard(T, it, x0 + i * (w + 0.22), 1.55, w, 5.4, {
        imageH: 2.45, titleSize: 13, descSize: 10, joinDesc: true, statsInRow: Math.min(Math.max(it.stats.length, 1), 5),
        statSize: it.stats.some((st) => st.value.length > 7) ? 12.5 : 18,
      })));
      return out;
    },
  },
  {
    id: 'feature', name: 'Single feature', when: (m) => m.items.length === 1,
    render(m, { theme: T }) {
      const it = m.items[0], out = header(m, T);
      out.push(...card(T, M, 1.55, CW, 5.4), image(T, M, 1.61, 6.3, 5.34, it.pic, it.pic && it.pic.graphic ? 'contain' : 'cover'));
      out.push(text(M + 6.5, 1.85, CW - 6.7, 2.6, [para(titleRuns(it, 16, T.onSurface), { after: 8 }), ...it.desc.map((d) => para(run(d, 11, T.soft), { after: 6 }))]));
      const inRow = Math.min(3, Math.max(it.stats.length, 1)), cw = (CW - 6.7) / inRow;
      out.push(rect(M + 6.5, 4.85, CW - 6.7, 0.01, T.line));
      it.stats.forEach((s, j) => out.push(text(M + 6.5 + (j % inRow) * cw, 5.0 + Math.floor(j / inRow) * 0.85, cw, 0.8, [
        para(run(s.value, 22, j === 0 ? ca(T) : T.onSurface, true)), para(run(s.label, 10, T.muted))])));
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
const axis = (T) => ({ textColor: T.chartText, gridColor: T.chartGrid, legendColor: T.onSurface, labelColor: T.onSurface });

function lineSpec(T, c, o = {}) {
  const series = [...c.series].sort((a, b) => (parseFloat(b.name) || 0) - (parseFloat(a.name) || 0)).reverse(); // current year last = highlight
  return {
    kind: 'line', cats: c.cats, series, colors: series.map((s, i, arr) => (i === arr.length - 1 ? T.chart[0] : T.chart[1])),
    title: o.title === false ? '' : c.title, legend: true, numFmt: numFmt(c.series.flatMap((s) => s.values)), ...axis(T),
  };
}
function lineStat(c) {
  const series = [...c.series].sort((a, b) => (parseFloat(b.name) || 0) - (parseFloat(a.name) || 0));
  const cur = series[0], prev = series[1], i = lastIdx(cur.values), comps = [];
  if (prev && i >= 0 && prev.values[i]) {
    const pct = (cur.values[i] / prev.values[i] - 1) * 100;
    comps.push(`${pct >= 0 ? '+' : ''}${pct.toFixed(0)}% on the same point in ${prev.name}`);
  }
  const label = c.title.replace(/\s*\d{2}\s*vs\s*\d{2}\s*$/i, '').trim() || c.title;
  return { num: fmt(cur.values[i]), label: `${label} to ${c.cats[i] || 'date'}`.trim(), comps };
}
function barSpec(T, c) {
  const vals = c.series[0] ? c.series[0].values : [];
  return {
    kind: 'bar', cats: c.cats, series: [c.series[0] || { name: '', values: [] }], pointColors: c.cats.map((_, i) => T.chart[i % T.chart.length]),
    title: c.title, legend: false, numFmt: numFmt(vals), dataLabels: true, ...axis(T),
  };
}
function donutSpec(T, c) {
  const pc = c.cats.map((k, i) => T.pillars[k.toLowerCase().trim()] || T.chart[i % T.chart.length]);
  return { kind: 'donut', cats: c.cats, series: [c.series[0] || { name: '', values: [] }], pointColors: pc, title: c.title, legend: true, dataLabels: true, ...axis(T) };
}
const specFor = (T, c) => (c.kind === 'line' ? lineSpec(T, c) : /pie|doughnut|donut/.test(c.kind) ? donutSpec(T, c) : barSpec(T, c));
const isRate = (c) => /rate|%|percent/i.test(c.title) || (c.series[0] && c.series[0].values.every((v) => v == null || v < 100) && c.series[0].values.some((v) => v % 1));
const leftCopy = (m) => (m.facts && m.facts.length) || (m.bullets && m.bullets.length) || (m.figures && m.figures.length);

const chartOptions = [
  {
    id: 'statcards', name: 'Stat cards above each chart', when: (m) => m.charts.every((c) => c.kind === 'line'),
    render(m, { theme: T }) {
      const out = header(m, T), list = m.charts.slice(0, 3), n = list.length, w = (CW - 0.22 * (n - 1)) / n;
      list.forEach((c, i) => {
        const x = M + i * (w + 0.22);
        out.push(...statCard(T, lineStat(c), x, 1.45, w, 1.75), ...card(T, x, 3.35, w, 3.45), chart(x, 3.45, w, 3.3, lineSpec(T, c, { title: false })));
      });
      return out;
    },
  },
  {
    id: 'grid', name: 'Charts on cards',
    render(m, { theme: T }) {
      const left = leftCopy(m), out = header(m, T), list = m.charts.slice(0, 4), n = list.length;
      const gx = left ? 4.95 : M, gw = left ? SW - M - 4.95 : CW;
      if (left) out.push(...factsAndBullets(m, T, M, TOP, 4.2, BOTTOM - TOP));
      const cols = n === 1 ? 1 : 2, rows = Math.ceil(n / cols), gap = 0.15;
      const w = (gw - gap * (cols - 1)) / cols, h = (BOTTOM - 1.45 - gap * (rows - 1)) / rows;
      list.forEach((c, i) => {
        let x = gx + (i % cols) * (w + gap); const y = 1.45 + Math.floor(i / cols) * (h + gap);
        if (n === 3 && i === 2 && !left) x = gx + (gw - w) / 2;
        out.push(...card(T, x, y, w, h), chart(x, y + 0.08, w, h - 0.12, specFor(T, c)));
      });
      return out;
    },
  },
  {
    id: 'insight', name: 'Big chart with insight panel',
    render(m, { theme: T }) {
      const out = header(m, T), c = m.charts[0];
      out.push(...card(T, M, 1.45, 8.2, 5.4), chart(M, 1.6, 8.2, 5.2, specFor(T, c)));
      const x = M + 8.45, w = SW - M - x;
      const ps = [para(run('INSIGHT', 10, T.accent, true), { after: 6 })];
      if (c.kind === 'line') { const s = lineStat(c); ps.push(para(run(s.num, 40, T.text, true)), para(run(s.label, 12, T.muted), { after: 6 }), ...s.comps.map((k) => para(run(k, 12, T.text), { after: 12 }))); }
      ps.push(...copyParas(m, T, 10.5));
      if (m.charts.length > 1) m.charts.slice(1, 3).forEach((k) => { if (k.kind === 'line') { const s = lineStat(k); ps.push(para([run(s.num + '  ', 16, T.text, true), run(s.label, 10, T.muted)], { after: 4 })); } });
      out.push(text(x, 1.55, w, 5.3, ps));
      return out;
    },
  },
  {
    id: 'hero', name: 'Hero chart with supporting charts', when: (m) => m.charts.length >= 2 && m.charts.length <= 3,
    render(m, { theme: T }) {
      const out = header(m, T), [first, ...rest] = m.charts;
      out.push(...card(T, M, 1.45, 7.9, 5.35), chart(M, 1.6, 7.9, 5.15, specFor(T, first)));
      const h = rest.length === 1 ? 5.35 : 2.6;
      rest.forEach((c, i) => out.push(...card(T, 8.53, 1.45 + i * (h + 0.15), 4.38, h), chart(8.53, 1.55 + i * (h + 0.15), 4.38, h - 0.15, specFor(T, c))));
      return out;
    },
  },
  {
    id: 'share', name: 'Share of total bars', when: (m) => m.charts.length >= 2 && m.charts.every((c) => c.kind === 'bar') && m.charts.some((c) => !isRate(c)),
    render(m, { theme: T }) {
      const out = header(m, T), left = leftCopy(m), gx = left ? 4.95 : M, gw = left ? SW - M - 4.95 : CW;
      if (left) out.push(...factsAndBullets(m, T, M, TOP, 4.2, BOTTOM - TOP));
      const shares = m.charts.filter((c) => !isRate(c)), rates = m.charts.filter(isRate), cats = shares[0].cats;
      const pal = [T.chart[0], T.chart[1], T.chart[2], T.chart[3], T.chart[4]];
      out.push(...card(T, gx, 1.35, gw, Math.min(1.25 * shares.length + 0.85, 4.1)));
      out.push(chart(gx + 0.1, 1.45, gw - 0.2, Math.min(1.25 * shares.length + 0.7, 3.95), {
        kind: 'bar100', cats: shares.map((c) => c.title.replace(/\s*per\s+\w+\s*$/i, '')),
        series: cats.map((k, i) => ({ name: k, values: shares.map((c) => (c.series[0] ? c.series[0].values[i] : null)) })),
        colors: cats.map((_, i) => pal[i % 5]), legend: true, dataLabels: true, numFmt: '[>=1000000]0.0,,"M";[>=1000]0.0,"K";0', ...axis(T),
      }));
      rates.slice(0, 1).forEach((c) => {
        const y = 5.6;
        out.push(text(gx, y, gw, 0.35, [para(run(c.title.replace(/\s*per\s+\w+\s*$/i, ''), 13, T.text, true))]));
        const n = c.cats.length, w = (gw - 0.12 * (n - 1)) / n;
        c.cats.forEach((k, i) => out.push(text(gx + i * (w + 0.12), y + 0.42, w, 0.95, [
          para(run(`${(+c.series[0].values[i] || 0).toFixed(1)}%`, 24, i === 0 ? T.statNum : T.onSurface, true)), para(run(k, 10, i === 0 ? T.onStat : T.muted))],
        { fill: i === 0 ? T.statFill : T.surface, inset: 0.14 })));
      });
      return out;
    },
  },
  {
    id: 'donutcard', name: 'Donut on a card with copy', when: (m) => m.charts.length === 1 && /pie|doughnut|donut/.test(m.charts[0].kind),
    render(m, { theme: T }) {
      const out = header(m, T);
      out.push(...factsAndBullets(m, T, M, 1.55, 6.9, BOTTOM - 1.55));
      out.push(...card(T, 7.6, 1.5, 5.31, 5.35), chart(7.7, 1.65, 5.11, 5.1, donutSpec(T, m.charts[0])));
      return out;
    },
  },
  {
    id: 'pillarcards', name: 'Share cards', when: (m) => m.charts.length === 1 && /pie|doughnut|donut/.test(m.charts[0].kind),
    render(m, { theme: T }) {
      const out = header(m, T), c = m.charts[0], vals = c.series[0].values.map((v) => v || 0), tot = vals.reduce((a, v) => a + v, 0) || 1;
      const n = c.cats.length, w = (CW - 0.2 * (n - 1)) / n;
      c.cats.forEach((k, i) => {
        const x = M + i * (w + 0.2), col = T.pillars[k.toLowerCase().trim()] || T.chart[i % T.chart.length];
        out.push(rect(x, 1.55, w, 1.75, T.surface), rect(x, 1.55, w, 0.08, col));
        out.push(text(x, 1.65, w, 1.65, [para(run(`${Math.round((vals[i] / tot) * 100)}%`, 30, T.onSurface, true)), para(run(k, 12, T.onSurface), { after: 3 }), para(run(`${fmt(vals[i])} posts`, 10, T.muted))], { inset: 0.18 }));
      });
      out.push(...factsAndBullets(m, T, M, 3.6, CW, BOTTOM - 3.6));
      return out;
    },
  },
];

// ================= learnings / text =================
const learnings = [
  {
    id: 'numbered', name: 'Numbered cards',
    render(m, { theme: T }) {
      const out = header(m, T), items = m.items.slice(0, 4), w = m.image ? 7.05 : CW;
      if (m.image) out.push(image(T, 7.73, 0, SW - 7.73, SH, m.image));
      const g = 0.12, h = (BOTTOM - 1.45 - g * (items.length - 1)) / items.length;
      items.forEach((it, i) => {
        const y = 1.45 + i * (h + g);
        out.push(...card(T, M, y, w, h), text(M, y + 0.06, 0.75, h - 0.06, [para(run(String(i + 1), 30, ca(T), true), { align: 'center' })], { inset: 0.12 }));
        out.push(text(M + 0.7, y + 0.06, w - 0.75, h - 0.06, [...(it.head ? [para(run(it.head, 12, ca(T), true), { after: 5 })] : []), para(run(it.body, it.body.length > 450 ? 9 : 9.5, T.onSurface))], { inset: 0.14 }));
      });
      return out;
    },
  },
  {
    id: 'insight', name: 'Insight statement',
    render(m, { theme: T }) {
      const [first, ...rest] = m.items, ink = T.dark ? 'FFFFFF' : T.text;
      const out = [rect(0, 0, SW, SH, T.fade)];
      if (m.image) out.push(...photoFade(T, 5.4, 0, SW - 5.4, SH, m.image, 'left', 0.65));
      out.push(text(M, 0.42, 7, 0.6, [para(run(m.title || '', 26, ink))]));
      out.push(text(M, 1.35, 6.8, 5.6, [
        para(run('INSIGHT', 10, T.accent, true), { after: 4 }),
        para(run(first.head || first.body, 24, ink), { after: 14 }),
        ...(first.head ? [para(run(first.body, 11, T.dark ? 'D9D9D9' : T.soft), { after: 12 })] : []),
        ...rest.slice(0, 4).flatMap((it) => [...(it.head ? [para(run(it.head, 10, T.accent, true), { after: 2 })] : []), para(run(it.body, 10, T.dark ? 'D9D9D9' : T.soft), { after: 8, bullet: it.head ? null : T.accent })]),
      ]));
      return out;
    },
  },
  {
    id: 'circles', name: 'Numbered circles', when: (m) => m.items.length >= 2 && m.items.length <= 4,
    render(m, { theme: T }) {
      const out = header(m, T), n = m.items.length, w = (CW - 0.3 * (n - 1)) / n;
      m.items.forEach((it, i) => {
        const x = M + i * (w + 0.3);
        out.push(rect(x, 2.2, w, 4.6, T.surface), rect(x, 2.2, w, 0.04, ca(T)));
        out.push(ellipse(x + w / 2 - 0.42, 1.78, 0.84, 0.84, ca(T)), text(x + w / 2 - 0.42, 1.78, 0.84, 0.84, [para(run(String(i + 1), 20, ca(T) === T.accent ? T.onAccent : 'FFFFFF', true), { align: 'center' })], { valign: 'middle' }));
        out.push(text(x, 2.75, w, 4.0, [...(it.head ? [para(run(it.head, 13, T.onSurface, true), { after: 8, align: 'center' })] : []), para(run(it.body, 10, T.soft), { align: 'center' })], { inset: 0.2 }));
      });
      return out;
    },
  },
  {
    id: 'columns', name: 'Side by side', when: (m) => m.items.length <= 3,
    render(m, { theme: T }) {
      const out = header(m, T), n = m.items.length, w = (CW - 0.27 * (n - 1)) / n;
      m.items.forEach((it, i) => {
        const x = M + i * (w + 0.27);
        out.push(...card(T, x, 1.55, w, 5.4), text(x, 1.7, w, 0.9, [para(run(pad2(i + 1), 36, ca(T), true))], { inset: 0.2 }));
        out.push(text(x, 2.6, w, 4.2, [...(it.head ? [para(run(it.head, 13, T.onSurface, true), { after: 8 })] : []), para(run(it.body, 10.5, T.soft))], { inset: 0.2 }));
      });
      return out;
    },
  },
  {
    id: 'panel', name: 'Headings and bullets',
    render(m, { theme: T }) {
      const out = header(m, T), ps = [];
      m.items.forEach((it) => { if (it.head) ps.push(para(run(it.head, 12.5, T.accent, true), { after: 5 })); ps.push(para(run(it.body, 11, T.text), { after: 14, bullet: T.accent })); });
      out.push(text(M, 1.55, m.image ? 7.0 : CW, 5.4, ps));
      if (m.image) out.push(image(T, 7.73, 0, SW - 7.73, SH, m.image));
      return out;
    },
  },
];

const textOpts = [
  {
    id: 'bullets', name: 'Tidy bullets',
    render(m, { theme: T }) {
      const out = header(m, T), w = m.image ? 7.0 : CW;
      out.push(...factsAndBullets(m, T, M, 1.55, w, BOTTOM - 1.55));
      if (m.image) out.push(...card(T, 7.75, 1.5, 5.16, 5.4), image(T, 7.75, 1.56, 5.16, 5.34, m.image));
      return out;
    },
  },
  {
    id: 'fade', name: 'Copy over photo fade', when: (m) => !!m.image,
    render(m, { theme: T }) {
      const ink = T.dark ? 'FFFFFF' : T.text;
      return [rect(0, 0, SW, SH, T.fade), ...photoFade(T, 5.0, 0, SW - 5.0, SH, m.image, 'left', 0.65),
        text(M, 0.42, 7, 0.6, [para(run(m.title || m.subtitle || '', 26, ink))]),
        text(M, 1.4, 6.6, 5.5, (m.bullets || []).map((b) => para(run(b, 11, T.dark ? 'D9D9D9' : T.soft), { after: 8, bullet: T.accent })))];
    },
  },
  {
    id: 'cards', name: 'Points as cards', when: (m) => m.bullets && m.bullets.length >= 2 && m.bullets.length <= 6,
    render(m, { theme: T }) {
      const out = header(m, T), b = m.bullets, cols = b.length <= 3 ? b.length : Math.ceil(b.length / 2), rows = Math.ceil(b.length / cols);
      const w = (CW - 0.2 * (cols - 1)) / cols, h = (BOTTOM - 1.55 - 0.2 * (rows - 1)) / rows;
      b.forEach((t, i) => {
        const x = M + (i % cols) * (w + 0.2), y = 1.55 + Math.floor(i / cols) * (h + 0.2);
        out.push(...card(T, x, y, w, h), text(x, y + 0.1, w, h - 0.1, [para(run(t, 11, T.onSurface), { bullet: ca(T) })], { inset: 0.2 }));
      });
      return out;
    },
  },
];

// ---------- reading results tables ----------
const METRIC = /impression|view|engagement|eng\.?\s*rate|reach|post|spend|cpv|cost|click|follower|share|like|comment|save/i;
const LOWER_IS_BETTER = /cpv|cost/i;
function cellNum(s) {
  if (s == null) return null;
  const t = String(s).replace(/[£$€,\s]/g, '');
  if (!/^-?\d+(\.\d+)?%?$/.test(t)) return null;
  return parseFloat(t);
}
function shortCell(s) {
  const t = String(s || '').trim();
  if (/%$/.test(t)) return t;
  const n = cellNum(t);
  if (n == null) return t || '–';
  const money = /^[£$€]/.test(t) ? t[0] : '';
  if (Math.abs(n) >= 10000) return money + fmt(n);
  return t;
}
// Turns table rows into { entities, metrics } when rows are metrics (Posts, Views…) and
// columns are pieces/channels; or { cats, cols } when rows are categories (18-24, UK…).
function readTable(rows) {
  if (!rows || rows.length < 2) return null;
  const head = rows[0].map((h) => h.trim());
  const body = rows.slice(1).filter((r) => r.some((c) => c && c.trim()));
  const metricRows = body.filter((r) => METRIC.test(r[0] || ''));
  if (metricRows.length >= 2 && head.length >= 2) {
    const entities = head.slice(1).map((name, i) => ({ name: name || `Column ${i + 1}`, total: /^totals?$/i.test(name) }));
    const metrics = metricRows.map((r) => ({ label: r[0].trim(), raw: r.slice(1), vals: r.slice(1).map(cellNum) }));
    return { kind: 'metrics', entities, metrics };
  }
  const numericCols = head.slice(1).map((h, i) => ({ name: h || `Value ${i + 1}`, vals: body.map((r) => cellNum(r[i + 1])) }))
    .filter((c) => c.vals.filter((v) => v != null).length >= 2);
  if (numericCols.length) return { kind: 'cats', cats: body.map((r) => r[0]), cols: numericCols };
  return null;
}
const PREF = [/view/i, /impression/i, /engagement$|engagements/i, /rate/i];
function pickMetrics(metrics, n) {
  const picked = [];
  for (const re of PREF) { const m = metrics.find((x) => re.test(x.label) && !picked.includes(x)); if (m) picked.push(m); }
  for (const m of metrics) if (!picked.includes(m)) picked.push(m);
  return picked.slice(0, n);
}
const firstMetricTable = (m) => m.tables.map(readTable).find((t) => t && t.kind === 'metrics');
const tableBox = (T, x, y, w, h, rows) => ({ t: 'table', x, y, w, h, rows, head: T.accent, onHead: T.onAccent, rowA: T.surface, rowB: T.bg, ink: T.onSurface });

function photosColumn(T, pics, x, y, w, h) {
  const list = pics.slice(0, 3), g = 0.15, ph = (h - g * (list.length - 1)) / Math.max(list.length, 1);
  return list.map((p, i) => image(T, x, y + i * (ph + g), w, ph, p, p.graphic ? 'contain' : 'cover'));
}

const tableOpts = [
  {
    id: 'scorecards', name: 'Scorecards',
    when: (m) => { const t = firstMetricTable(m); return t && t.entities.length >= 2 && t.entities.length <= 5; },
    render(m, { theme: T }) {
      const out = header(m, T), t = firstMetricTable(m), n = t.entities.length;
      const hasCopy = m.bullets && m.bullets.length, cardH = hasCopy ? 3.35 : 5.35, gap = 0.18, w = (CW - gap * (n - 1)) / n;
      const [lead, ...others] = pickMetrics(t.metrics, 7);
      // best value in each row (ignoring a Totals column)
      const best = t.metrics.map((mt) => {
        const vals = mt.vals.map((v, i) => (t.entities[i].total ? null : v));
        const ok = vals.filter((v) => v != null);
        if (ok.length < 2) return -1;
        const target = LOWER_IS_BETTER.test(mt.label) ? Math.min(...ok) : Math.max(...ok);
        return vals.indexOf(target);
      });
      t.entities.forEach((e, i) => {
        const x = M + i * (w + gap), y = 1.5, totalCard = e.total;
        const fill = totalCard ? T.statFill : T.surface, ink = totalCard ? T.onStat : T.onSurface, big = totalCard ? T.statNum : T.onSurface;
        out.push(rect(x, y, w, cardH, fill), rect(x, y, w, 0.06, totalCard ? T.bar : ca(T)));
        out.push(text(x, y + 0.1, w, 0.75, [para(run(e.name, 13, ink, true))], { inset: 0.16, valign: 'middle' }));
        out.push(text(x, y + 0.85, w, 0.85, [para(run(shortCell(lead.raw[i]), 30, big, true)), para(run(lead.label, 10, totalCard ? T.onStat : T.muted))], { inset: 0.16 }));
        const rowsN = Math.ceil(others.length / 2), rh = Math.min(0.55, (cardH - 1.85) / Math.max(rowsN, 1));
        others.forEach((mt, j) => {
          const isBest = best[t.metrics.indexOf(mt)] === i;
          out.push(text(x + 0.16 + (j % 2) * ((w - 0.32) / 2), y + 1.8 + Math.floor(j / 2) * rh, (w - 0.32) / 2, rh, [
            para(run(shortCell(mt.raw[i]), 13, isBest && !totalCard ? ca(T) : ink, true)), para(run(mt.label, 8.5, totalCard ? T.onStat : T.muted)),
          ]));
        });
      });
      if (hasCopy) {
        const bl = m.bullets.filter((b) => !/:$/.test(b)).slice(0, 4), k = Math.max(Math.min(bl.length, 2), 1), bw = (CW - 0.4 * (k - 1)) / k;
        const halves = k === 1 ? [bl] : [bl.slice(0, Math.ceil(bl.length / 2)), bl.slice(Math.ceil(bl.length / 2))];
        halves.forEach((h, i) => out.push(text(M + i * (bw + 0.4), 5.05, bw, 1.95, h.map((b) => para(run(b, fitSize(h, bw - 0.25, 1.9, 9.5, 7), T.text), { after: 6, bullet: T.accent })))));
      }
      return out;
    },
  },
  {
    id: 'tiles', name: 'Key results with photos',
    when: (m) => !!firstMetricTable(m),
    render(m, { theme: T }) {
      const out = header(m, T), t = firstMetricTable(m);
      let col = t.entities.findIndex((e) => e.total);
      if (col < 0) col = t.entities.length - 1;
      const metrics = pickMetrics(t.metrics, 6).filter((mt) => mt.raw[col] && mt.raw[col].trim() && mt.raw[col].trim() !== '/');
      const pics = m.pics || [];
      const copyW = 4.1, picsW = pics.length ? 3.4 : 0;
      const tx = M + copyW + 0.3, tw = SW - M - tx - (picsW ? picsW + 0.3 : 0);
      out.push(text(M, 1.5, copyW, 5.45, [
        para(run(t.entities[col].total ? 'Totals' : t.entities[col].name, 11, T.accent, true), { after: 6 }),
        ...((bl) => bl.map((b) => para(run(b, fitSize(bl, copyW, 5.0, 10), T.text), { after: 6 })))((m.bullets || []).filter((b) => !/^(overview|evaluation of content)$/i.test(b)).slice(0, 4)),
      ]));
      const cols = 2, rows = Math.ceil(metrics.length / cols), g = 0.16, w = (tw - g) / cols, h = Math.min(1.6, (5.45 - g * (rows - 1)) / Math.max(rows, 1));
      metrics.forEach((mt, i) => out.push(...statCard(T, { num: shortCell(mt.raw[col]), label: mt.label, comps: [] }, tx + (i % cols) * (w + g), 1.5 + Math.floor(i / cols) * (h + g), w, h, { big: 26 })));
      if (picsW) out.push(...photosColumn(T, pics, SW - M - picsW, 1.5, picsW, 5.45));
      return out;
    },
  },
  {
    id: 'charts', name: 'Table as charts',
    when: (m) => m.tables.map(readTable).some((t) => t && (t.kind === 'cats' || t.entities.filter((e) => !e.total).length >= 2)),
    render(m, { theme: T }) {
      const out = header(m, T), specs = [];
      for (const t of m.tables.map(readTable)) {
        if (!t) continue;
        if (t.kind === 'metrics') {
          const keep = t.entities.map((e, i) => (e.total ? -1 : i)).filter((i) => i >= 0);
          for (const mt of pickMetrics(t.metrics, 3)) {
            if (keep.filter((i) => mt.vals[i] != null).length < 2) continue;
            specs.push({ title: mt.label, cats: keep.map((i) => t.entities[i].name), values: keep.map((i) => mt.vals[i] ?? 0), horizontal: false });
          }
        } else {
          for (const c of t.cols) specs.push({ title: c.name, cats: t.cats.slice(0, 12), values: c.vals.slice(0, 12).map((v) => v ?? 0), horizontal: true });
        }
      }
      const list = specs.slice(0, 3), left = withCopy(m), gx = left ? 4.6 : M, gw = SW - M - gx;
      if (left) out.push(...factsAndBullets(m, T, M, TOP, 3.9, BOTTOM - TOP));
      const n = Math.max(list.length, 1), w = (gw - 0.18 * (n - 1)) / n;
      list.forEach((s, i) => {
        const x = gx + i * (w + 0.18);
        const spec = { ...barSpec(T, { title: s.title, cats: s.cats, series: [{ name: s.title, values: s.values }] }), horizontal: s.horizontal };
        if (s.horizontal) { const top = Math.max(...s.values); spec.pointColors = s.values.map((v) => (v === top ? T.chart[0] : T.chart[1])); }
        out.push(...card(T, x, 1.45, w, 5.5), chart(x, 1.53, w, 5.38, spec));
      });
      return out;
    },
  },
  {
    id: 'table', name: 'Styled table with copy',
    render(m, { theme: T }) {
      const out = header(m, T), left = withCopy(m), pics = m.pics || [];
      if (left) out.push(...factsAndBullets(m, T, M, TOP, 3.9, BOTTOM - TOP));
      const picsW = pics.length ? 3.0 : 0;
      const gx = left ? 4.6 : M, gw = SW - M - gx - (picsW ? picsW + 0.25 : 0), n = Math.min(m.tables.length, 3), w = (gw - 0.2 * (n - 1)) / n;
      m.tables.slice(0, 3).forEach((rows, i) => out.push(tableBox(T, gx + i * (w + 0.2), TOP, w, BOTTOM - TOP, rows.map((r) => r.map(shortCell)))));
      if (picsW) out.push(...photosColumn(T, pics, SW - M - picsW, TOP, picsW, BOTTOM - TOP));
      return out;
    },
  },
];

export const OPTIONS = { cover, divider: dividers, contents, stats, cards, charts: chartOptions, learnings, text: textOpts, table: tableOpts };

export function optionsFor(model) {
  return (OPTIONS[model.type] || textOpts).filter((o) => !o.when || o.when(model));
}

// every layout starts from the theme background
export function renderWith(option, model, ctx) {
  return [rect(0, 0, SW, SH, ctx.theme.bg), ...option.render(model, ctx)];
}
