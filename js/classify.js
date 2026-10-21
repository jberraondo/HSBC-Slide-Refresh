// Works out what kind of slide each parsed slide is, and pulls out the pieces
// the templates need. Every extractor works from the same parsed slide, so the
// person can override the detected type in the page.

export const TYPES = {
  cover: 'Cover',
  divider: 'Section divider',
  contents: 'Contents',
  stats: 'Headline figures',
  cards: 'Content cards',
  charts: 'Charts',
  learnings: 'Learnings / text',
  table: 'Table',
  text: 'Text',
};

const NUM = /^(?:[£$€]?\d[\d,]*(?:\.\d+)?\s?(?:[MKBk]\b)?%?|x)(?=\s|$)/;
const area = (e) => (e.w || 0) * (e.h || 0);
const words = (s) => (s.match(/\S+/g) || []).length;
const SKIP_PH = new Set(['sldNum', 'ftr', 'dt']);

function prep(slide) {
  const els = slide.elements.filter((e) => e.x != null && e.x < 13.2 && e.x + e.w > 0.1 && e.y < 7.4 && e.y + e.h > 0.05);
  // owner tags like "Jacob" sit in the top right
  const isTag = (e) => e.kind === 'text' && e.y < 0.7 && e.x > 5 && e.paras.length === 1 && e.paras[0].text.length < 16 && !(e.ph && /title/i.test(e.ph.type));
  const isLogo = (e) => e.kind === 'pic' && (area(e) < 0.6 || /logo/i.test(e.name));
  const texts = els.filter((e) => e.kind === 'text' && e.paras.length && !(e.ph && SKIP_PH.has(e.ph.type)) && !isTag(e)
    && !/^internal$/i.test(e.paras.map((p) => p.text).join(' ')));
  const pics = els.filter((e) => e.kind === 'pic' && !isLogo(e));
  const charts = els.filter((e) => e.kind === 'chart');
  const tables = els.filter((e) => e.kind === 'table');
  // title and subtitle
  let title = texts.find((e) => e.ph && /title/i.test(e.ph.type) && e.ph.type !== 'subTitle');
  if (!title) title = texts.filter((e) => e.y < 1.2).sort((a, b) => a.y - b.y)[0];
  let sub = texts.find((e) => e !== title && e.ph && e.ph.type === 'subTitle')
    || texts.filter((e) => e !== title && e.y < 1.7 && e.y + e.h < 2.3 && e.paras.length === 1 && e.paras[0].text.length < 170)
      .sort((a, b) => a.y - b.y)[0];
  const body = texts.filter((e) => e !== title && e !== sub);
  return {
    title: title ? title.paras.map((p) => p.text).join(' ') : '',
    subtitle: sub ? sub.paras.map((p) => p.text).join(' ') : '',
    body, pics, charts, tables,
  };
}

const allParas = (blocks) => blocks.flatMap((b) => b.paras);

// ---------- extractors ----------

function figuresFrom(blocks) {
  const figures = [], rest = [];
  for (const b of blocks) {
    const p = b.paras;
    const m = p[0].text.match(NUM);
    const isList = p.length >= 2 && p.every((x) => NUM.test(x.text)) && !p.slice(1).some((x) => /^[+\-=]/.test(x.text));
    if (m && p[0].text.length < 40 && !isList) {
      let num = m[0].trim(), label = p[0].text.slice(m[0].length).trim(), comps = p.slice(1).map((x) => x.text);
      if (!label && comps.length && !NUM.test(comps[0])) label = comps.shift();
      if (label && label.length < 40) { figures.push({ num, label, comps: comps.filter((c) => c.length < 90), y: b.y, x: b.x }); continue; }
    }
    rest.push(b);
  }
  figures.sort((a, b) => (Math.abs(a.y - b.y) < 0.4 ? a.x - b.x : a.y - b.y));
  return { figures, rest };
}

function splitCopy(blocks) {
  // short lines (e.g. "82 Posts Q3 '26") are facts, longer paragraphs are bullets
  const facts = [], bullets = [];
  for (const b of blocks.sort((a, c) => a.y - c.y || a.x - c.x)) {
    const short = b.paras.every((p) => p.text.length < 50);
    for (const p of b.paras) (short ? facts : bullets).push(p.text);
  }
  return { facts, bullets };
}

function statLine(text) {
  let m = text.match(/^([^:]{2,30}):\s*(.+)$/);
  if (m) return { value: m[2].trim(), label: m[1].trim() };
  m = text.match(/^([£$€]?[\d.,]+\s?[MKBk]?%?)\s+(.{2,30})$/);
  if (m) return { value: m[1].replace(/\s+(?=[MKBk]\b)/, ''), label: m[2].trim() };
  return null;
}

function cardsFrom(p) {
  const pics = p.pics.filter((e) => area(e) >= 1.2 && e.y > 0.9).sort((a, b) => a.x - b.x);
  const used = new Set();
  const items = [];
  for (const pic of pics) {
    const cx = pic.x + pic.w / 2;
    const col = p.body.filter((b) => !used.has(b) && Math.abs(b.x + b.w / 2 - cx) < Math.max(1.4, pic.w * 0.6) && b.y >= pic.y + pic.h * 0.5)
      .sort((a, b) => a.y - b.y || a.x - b.x);
    col.forEach((b) => used.add(b));
    const paras = allParas(col);
    if (!paras.length) { items.push({ pic, titleRuns: [], titleText: '', desc: [], stats: [] }); continue; }
    const title = paras[0];
    const desc = [], stats = [];
    let inStats = false;
    for (let i = 1; i < paras.length; i++) {
      const t = paras[i].text;
      if (/^key results$/i.test(t)) { inStats = true; continue; }
      const sl = statLine(t);
      // value then label on separate lines (already restyled slides)
      if (NUM.test(t) && t.length < 16 && paras[i + 1] && paras[i + 1].text.length < 28 && !NUM.test(paras[i + 1].text)) {
        stats.push({ value: t, label: paras[i + 1].text }); i++; inStats = true; continue;
      }
      if (sl && (inStats || /\d/.test(sl.value))) { stats.push(sl); inStats = true; continue; }
      if (!inStats) desc.push(t);
    }
    items.push({ pic, titleRuns: title.runs, titleText: title.text, desc, stats });
  }
  const leftover = p.body.filter((b) => !used.has(b));
  return { items, ...splitCopy(leftover) };
}

function chartsFrom(p) {
  const charts = p.charts.sort((a, b) => (Math.abs(a.y - b.y) < 0.5 ? a.x - b.x : a.y - b.y)).map((c) => c.chart);
  const { figures, rest } = figuresFrom(p.body);
  return { charts, figures, ...splitCopy(rest) };
}

function learningsFrom(p) {
  const paras = allParas(p.body.sort((a, b) => a.y - b.y || a.x - b.x)).map((x) => x.text).filter((t) => !/^\d{1,2}$/.test(t));
  const items = [];
  for (let i = 0; i < paras.length; i++) {
    const t = paras[i];
    if (t.length < 110 && paras[i + 1] && paras[i + 1].length > t.length) { items.push({ head: t, body: paras[i + 1] }); i++; }
    else items.push({ head: '', body: t });
  }
  return { items, image: p.pics.sort((a, b) => area(b) - area(a))[0] || null };
}

function dividerFrom(p) {
  const all = allParas(p.body).map((x) => x.text).concat(p.title ? [p.title] : []);
  const cands = all.filter((t) => !/^\d{1,2}$/.test(t) && !/\|/.test(t) && t.length > 1);
  const title = cands.sort((a, b) => b.length - a.length)[0] || p.title || 'Section';
  return { title, image: p.pics.sort((a, b) => area(b) - area(a))[0] || null };
}

function contentsFrom(p) {
  const items = allParas(p.body.sort((a, b) => a.y - b.y)).map((x) => x.text.replace(/^\d{1,2}\s+/, ''))
    .filter((t) => !/^contents$/i.test(t) && !/^\d{1,2}$/.test(t) && !/\|/.test(t));
  return { items, image: p.pics.sort((a, b) => area(b) - area(a))[0] || null };
}

// ---------- detection ----------

export function detect(slide) {
  const p = prep(slide);
  const paraTexts = allParas(p.body).map((x) => x.text);
  const totalWords = paraTexts.reduce((n, t) => n + words(t), 0) + words(p.title);
  const bigPic = p.pics.some((e) => area(e) > 18);
  const ctr = slide.elements.some((e) => e.ph && e.ph.type === 'ctrTitle');
  if ((ctr || slide.number === 1) && totalWords <= 12 && !p.charts.length) return 'cover';
  if (/^contents$/i.test(p.title) || paraTexts.some((t) => /^contents$/i.test(t))) return 'contents';
  if (totalWords <= 10 && (bigPic || p.body.length + (p.title ? 1 : 0) <= 2) && !p.charts.length) return 'divider';
  if (p.charts.length) return 'charts';
  const cards = cardsFrom(p);
  if (cards.items.length && cards.items.length <= 8 && cards.items.every((i) => i.titleText && (i.stats.length || i.desc.length))) return 'cards';
  if (p.tables.length) return 'table';
  if (figuresFrom(p.body).figures.length >= 3) return 'stats';
  const learn = learningsFrom(p);
  if (learn.items.filter((i) => i.head).length >= 2) return 'learnings';
  return 'text';
}

export function extract(slide, type) {
  const p = prep(slide);
  const base = { type, title: p.title, subtitle: p.subtitle, number: slide.number };
  switch (type) {
    case 'divider': return { ...base, ...dividerFrom(p) };
    case 'cover': {
      // the biggest text is the title; skip boilerplate like "proprietary and confidential"
      const d = dividerFrom(p);
      const paras = slide.elements.filter((e) => e.kind === 'text').flatMap((e) => e.paras.map((x) => ({ ...x, ph: e.ph })))
        .filter((x) => x.text.length > 1 && !/confidential|proprietary|internal|^\d+$/i.test(x.text));
      const score = (x) => (x.size || 0) + (x.ph && /title/i.test(x.ph.type) ? 100 : 0);
      const sorted = [...paras].sort((a, b) => score(b) - score(a));
      const title = sorted[0] ? sorted[0].text : d.title;
      const subtitle = paras.filter((x) => x.text !== title).map((x) => x.text).slice(0, 2).join(' ');
      return { ...base, title, subtitle, image: d.image };
    }
    case 'contents': return { ...base, ...contentsFrom(p) };
    case 'stats': { const { figures, rest } = figuresFrom(p.body); return { ...base, figures, ...splitCopy(rest) }; }
    case 'cards': return { ...base, ...cardsFrom(p) };
    case 'charts': return { ...base, ...chartsFrom(p) };
    case 'learnings': return { ...base, ...learningsFrom(p) };
    case 'table': return { ...base, tables: p.tables.map((t) => t.rows), ...splitCopy(p.body) };
    default: return { ...base, ...splitCopy(p.body), image: p.pics.sort((a, b) => area(b) - area(a))[0] || null };
  }
}
