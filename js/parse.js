// Reads a .pptx into a plain model: slides -> elements (text, pictures, charts, tables).
// Works in the browser and in Node. Pass in the JSZip constructor.

const EMU = 914400;

const decode = (s) => s
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n))
  .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCharCode(parseInt(n, 16)))
  .replace(/&amp;/g, '&');

function attrs(tag) {
  const out = {};
  for (const m of tag.matchAll(/([\w:]+)="([^"]*)"/g)) out[m[1]] = m[2];
  return out;
}

function dirOf(path) { return path.slice(0, path.lastIndexOf('/') + 1); }

function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (base + target).split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}

function parseRels(xml, partPath) {
  const rels = {};
  if (!xml) return rels;
  const base = dirOf(partPath);
  for (const m of xml.matchAll(/<Relationship\b[^>]*\/?>/g)) {
    const a = attrs(m[0]);
    const external = a.TargetMode === 'External';
    rels[a.Id] = { type: a.Type || '', target: external ? a.Target : resolvePath(base, a.Target), external };
  }
  return rels;
}

// Returns the top-level child elements of `xml` whose tag is one of `tags`,
// balancing nested elements with the same tag (needed for groups).
function topLevel(xml, tags) {
  const out = [];
  const re = new RegExp(`<(${tags.map((t) => t.replace(':', '\\:')).join('|')})(?=[\\s>/])`, 'g');
  let i = 0;
  while (true) {
    re.lastIndex = i;
    const m = re.exec(xml);
    if (!m) break;
    const tag = m[1];
    const start = m.index;
    const open = new RegExp(`<${tag}(?=[\\s>/])`, 'g');
    const close = new RegExp(`</${tag}>`, 'g');
    // self-closing
    const tagEnd = xml.indexOf('>', start);
    if (xml[tagEnd - 1] === '/') { out.push({ tag, xml: xml.slice(start, tagEnd + 1) }); i = tagEnd + 1; continue; }
    let depth = 1, pos = tagEnd + 1;
    while (depth > 0) {
      open.lastIndex = pos; close.lastIndex = pos;
      const o = open.exec(xml), c = close.exec(xml);
      if (!c) { pos = xml.length; break; }
      if (o && o.index < c.index) {
        const e = xml.indexOf('>', o.index);
        if (xml[e - 1] !== '/') depth++;
        pos = e + 1;
      } else { depth--; pos = c.index + c[0].length; }
    }
    out.push({ tag, xml: xml.slice(start, pos) });
    i = pos;
  }
  return out;
}

function unwrapAlternateContent(xml) {
  return xml.replace(/<mc:AlternateContent\b[^>]*>[\s\S]*?<mc:Choice\b[^>]*>([\s\S]*?)<\/mc:Choice>[\s\S]*?<\/mc:AlternateContent>/g, '$1');
}

function xfrmOf(xml) {
  const m = xml.match(/<a:off x="(-?\d+)" y="(-?\d+)"\s*\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\s*\/>/);
  if (!m) return null;
  return { x: +m[1] / EMU, y: +m[2] / EMU, w: +m[3] / EMU, h: +m[4] / EMU };
}

function phOf(xml) {
  const m = xml.match(/<p:ph\b[^>]*\/?>/);
  if (!m) return null;
  const a = attrs(m[0]);
  return { type: a.type || 'body', idx: a.idx || '0' };
}

function paragraphs(xml, rels) {
  const body = xml.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/) || xml.match(/<a:txBody>([\s\S]*?)<\/a:txBody>/);
  if (!body) return [];
  const out = [];
  const bodyXml = body[1].replace(/<a:br\b[^>]*\/>|<a:br\b[^>]*>[\s\S]*?<\/a:br>/g, '<a:r><a:t> </a:t></a:r>');
  for (const p of bodyXml.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)) {
    const runs = [];
    let maxSize = 0, bold = false;
    for (const r of p[1].matchAll(/<a:(r|fld)\b[^>]*>([\s\S]*?)<\/a:\1>/g)) {
      const t = r[2].match(/<a:t>([\s\S]*?)<\/a:t>/);
      if (!t) continue;
      const rPr = (r[2].match(/<a:rPr\b[^>]*>/) || [''])[0];
      const ra = attrs(rPr);
      if (ra.sz) maxSize = Math.max(maxSize, +ra.sz / 100);
      if (ra.b === '1') bold = true;
      const hl = r[2].match(/<a:hlinkClick\b[^>]*r:id="(rId\d+)"/);
      const url = hl && rels[hl[1]] && rels[hl[1]].external ? rels[hl[1]].target : null;
      runs.push({ text: decode(t[1]).replace(/\u200b/g, ''), url });
    }
    const text = runs.map((r) => r.text).join('').replace(/\s+/g, ' ').trim();
    if (text) out.push({ text, runs, size: maxSize || null, bold });
  }
  return out;
}

async function parseChart(zip, path) {
  const xml = await zip.file(path)?.async('string');
  if (!xml) return null;
  const tm = xml.match(/<c:(lineChart|line3DChart|barChart|bar3DChart|doughnutChart|pieChart|pie3DChart|areaChart)>/);
  if (!tm) return null;
  const kind = tm[1].replace('3D', '').replace('Chart', ''); // line | bar | doughnut | pie | area
  const barDir = (xml.match(/<c:barDir val="(\w+)"/) || [])[1] || null;
  const grouping = (xml.match(/<c:grouping val="(\w+)"/) || [])[1] || null;
  const titleXml = (xml.match(/<c:title>([\s\S]*?)<\/c:title>/) || [])[1] || '';
  const title = decode([...titleXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]).join('')
    || (titleXml.match(/<c:v>([\s\S]*?)<\/c:v>/) || [])[1] || '').trim();
  const series = [];
  let cats = [];
  for (const s of xml.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)) {
    const sx = s[1];
    const name = decode(((sx.match(/<c:tx>([\s\S]*?)<\/c:tx>/) || [])[1] || '').match(/<c:v>([\s\S]*?)<\/c:v>/)?.[1] || '');
    const catXml = (sx.match(/<c:cat>([\s\S]*?)<\/c:cat>/) || [])[1] || '';
    const catCount = +((catXml.match(/<c:ptCount val="(\d+)"/) || [])[1] || 0);
    const c = new Array(catCount).fill('');
    for (const p of catXml.matchAll(/<c:pt idx="(\d+)">\s*<c:v>([\s\S]*?)<\/c:v>/g)) c[+p[1]] = decode(p[2]).trim();
    if (c.length > cats.length) cats = c;
    const valXml = (sx.match(/<c:val>([\s\S]*?)<\/c:val>/) || [])[1] || '';
    const n = +((valXml.match(/<c:ptCount val="(\d+)"/) || [])[1] || catCount);
    const v = new Array(n).fill(null);
    for (const p of valXml.matchAll(/<c:pt idx="(\d+)"[^>]*>\s*<c:v>([\s\S]*?)<\/c:v>/g)) v[+p[1]] = parseFloat(p[2]);
    series.push({ name, values: v });
  }
  return { kind, barDir, grouping, title, cats, series };
}

function tableOf(xml) {
  const rows = [];
  for (const tr of xml.matchAll(/<a:tr\b[^>]*>([\s\S]*?)<\/a:tr>/g)) {
    const cells = [];
    for (const tc of tr[1].matchAll(/<a:tc\b[^>]*>([\s\S]*?)<\/a:tc>/g)) {
      cells.push(decode([...tc[1].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => m[1]).join(' ')).replace(/\s+/g, ' ').trim());
    }
    rows.push(cells);
  }
  return rows;
}

async function placeholderMap(zip, path) {
  // positions of placeholders on a layout or master, keyed by "type:idx", "type:" and ":idx"
  const map = {};
  const xml = await zip.file(path)?.async('string');
  if (!xml) return map;
  for (const sp of topLevel(xml, ['p:sp'])) {
    const ph = phOf(sp.xml), xf = xfrmOf(sp.xml);
    if (!ph || !xf) continue;
    map[`${ph.type}:${ph.idx}`] = xf;
    if (!map[`${ph.type}:`]) map[`${ph.type}:`] = xf;
    if (!map[`:${ph.idx}`]) map[`:${ph.idx}`] = xf;
  }
  return map;
}

function lookupPh(ph, maps) {
  for (const m of maps) {
    const hit = m[`${ph.type}:${ph.idx}`] || m[`:${ph.idx}`] || m[`${ph.type}:`]
      || (ph.type === 'body' ? null : m[`${ph.type === 'ctrTitle' ? 'title' : ph.type}:`]);
    if (hit) return hit;
  }
  return null;
}

async function walk(zip, xml, rels, maps, transform, out) {
  const kids = topLevel(xml, ['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp']);
  for (const k of kids) {
    const name = decode((k.xml.match(/<p:cNvPr\b[^>]*name="([^"]*)"/) || [])[1] || '');
    if (k.tag === 'p:grpSp') {
      const g = k.xml.match(/<p:grpSpPr>([\s\S]*?)<\/p:grpSpPr>/);
      const gx = g && g[1].match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\/>\s*<a:chOff x="(-?\d+)" y="(-?\d+)"\/>\s*<a:chExt cx="(\d+)" cy="(\d+)"\/>/);
      let t = transform;
      if (gx) {
        const [ox, oy, ex, ey, cx, cy, cex, cey] = gx.slice(1).map(Number);
        const sx = cex ? ex / cex : 1, sy = cey ? ey / cey : 1;
        const inner = (p) => ({ x: ox / EMU + (p.x - cx / EMU) * sx, y: oy / EMU + (p.y - cy / EMU) * sy, w: p.w * sx, h: p.h * sy });
        t = (p) => transform(inner(p));
      }
      const innerXml = k.xml.replace(/^<p:grpSp\b[^>]*>/, '').replace(/<\/p:grpSp>$/, '').replace(/<p:nvGrpSpPr>[\s\S]*?<\/p:nvGrpSpPr>/, '').replace(/<p:grpSpPr>[\s\S]*?<\/p:grpSpPr>/, '');
      await walk(zip, innerXml, rels, maps, t, out);
      continue;
    }
    if (k.tag === 'p:cxnSp') continue;
    const ph = phOf(k.xml);
    let box = xfrmOf(k.tag === 'p:graphicFrame' ? (k.xml.match(/<p:xfrm>[\s\S]*?<\/p:xfrm>/) || [''])[0] : (k.xml.match(/<a:xfrm\b[\s\S]*?<\/a:xfrm>/) || [''])[0]);
    if (!box && ph) box = lookupPh(ph, maps);
    if (box) box = transform(box);
    const base = { name, ph, ...(box || { x: null, y: null, w: null, h: null }) };
    if (k.tag === 'p:sp') {
      const paras = paragraphs(k.xml, rels);
      const fill = (k.xml.match(/<p:spPr>[\s\S]*?<a:solidFill>\s*<a:srgbClr val="(\w{6})"/) || [])[1] || null;
      const blip = k.xml.match(/<a:blip\b[^>]*r:embed="(rId\d+)"/);
      if (blip && rels[blip[1]]) out.push({ ...base, kind: 'pic', image: rels[blip[1]].target, crop: null, paras });
      out.push({ ...base, kind: 'text', paras, fill });
    } else if (k.tag === 'p:pic') {
      const emb = k.xml.match(/<a:blip\b[^>]*r:embed="(rId\d+)"/);
      if (!emb || !rels[emb[1]] || rels[emb[1]].external) continue;
      const sr = k.xml.match(/<a:srcRect\b([^>]*)\/>/);
      const c = sr ? attrs(sr[1]) : {};
      const crop = { l: (+c.l || 0) / 100000, t: (+c.t || 0) / 100000, r: (+c.r || 0) / 100000, b: (+c.b || 0) / 100000 };
      out.push({ ...base, kind: 'pic', image: rels[emb[1]].target, crop });
    } else if (k.tag === 'p:graphicFrame') {
      const ch = k.xml.match(/<c:chart\b[^>]*r:id="(rId\d+)"/);
      if (ch && rels[ch[1]]) {
        const chart = await parseChart(zip, rels[ch[1]].target);
        if (chart) out.push({ ...base, kind: 'chart', chart });
      } else if (k.xml.includes('<a:tbl>')) {
        out.push({ ...base, kind: 'table', rows: tableOf(k.xml) });
      }
    }
  }
}

export async function parsePptx(data, JSZip) {
  const zip = await JSZip.loadAsync(data);
  const read = (p) => zip.file(p)?.async('string');
  const pres = await read('ppt/presentation.xml');
  if (!pres) throw new Error('This file does not look like a PowerPoint (.pptx) file.');
  const presRels = parseRels(await read('ppt/_rels/presentation.xml.rels'), 'ppt/presentation.xml');
  const sz = attrs((pres.match(/<p:sldSz\b[^>]*>/) || [''])[0]);
  const size = { w: (+sz.cx || 12192000) / EMU, h: (+sz.cy || 6858000) / EMU };
  const ids = [...pres.matchAll(/<p:sldId\b[^>]*r:id="(rId\d+)"/g)].map((m) => m[1]);
  const layoutCache = {};
  const slides = [];
  for (let n = 0; n < ids.length; n++) {
    const rel = presRels[ids[n]];
    if (!rel) continue;
    const path = rel.target;
    let xml = await read(path);
    if (!xml) continue;
    xml = unwrapAlternateContent(xml);
    const rels = parseRels(await read(dirOf(path) + '_rels/' + path.split('/').pop() + '.rels'), path);
    const layoutRel = Object.values(rels).find((r) => r.type.endsWith('/slideLayout'));
    let maps = [];
    if (layoutRel) {
      if (!layoutCache[layoutRel.target]) {
        const lrels = parseRels(await read(dirOf(layoutRel.target) + '_rels/' + layoutRel.target.split('/').pop() + '.rels'), layoutRel.target);
        const masterRel = Object.values(lrels).find((r) => r.type.endsWith('/slideMaster'));
        layoutCache[layoutRel.target] = [await placeholderMap(zip, layoutRel.target), masterRel ? await placeholderMap(zip, masterRel.target) : {}];
      }
      maps = layoutCache[layoutRel.target];
    }
    const tree = (xml.match(/<p:spTree>([\s\S]*)<\/p:spTree>/) || [])[1] || '';
    const elements = [];
    await walk(zip, tree, rels, maps, (p) => p, elements);
    // scale everything to a 13.333 x 7.5 inch slide if the deck uses another size
    const fx = 13.333 / size.w, fy = 7.5 / size.h;
    for (const e of elements) if (e.x != null) { e.x *= fx; e.w *= fx; e.y *= fy; e.h *= fy; }
    slides.push({ number: n + 1, hidden: /<p:sld\b[^>]*show="0"/.test(xml), elements });
  }
  const getImage = async (p) => {
    const f = zip.file(p);
    return f ? f.async('uint8array') : null;
  };
  return { size, slides, getImage };
}
