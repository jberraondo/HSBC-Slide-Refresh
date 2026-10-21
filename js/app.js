import { STYLE, THEMES, DEFAULT_THEME } from './style.js';
import { parsePptx } from './parse.js';
import { detect, extract, TYPES } from './classify.js';
import { optionsFor, renderWith } from './templates.js';
import { buildPptx, finishFile } from './build.js';
import { renderPreview } from './preview.js';
import { imageMaker, thumbFor, markGraphics, makeGradient } from './images.js';

const $ = (s) => document.querySelector(s);
const state = { deck: null, fileName: '', rows: [], thumbs: new Map(), showHidden: false };

// ---------- loading ----------
async function load(file) {
  setStatus(`Reading ${file.name}…`);
  try {
    const buf = await file.arrayBuffer();
    state.deck = await parsePptx(buf, window.JSZip);
    setStatus('Looking at the images…');
    await markGraphics(state.deck);
    state.fileName = file.name.replace(/\.pptx$/i, '');
    state.thumbs = new Map();
    state.rows = state.deck.slides.map((s) => {
      const type = detect(s);
      return makeRow(s, type, !s.hidden);
    });
    const cover = state.rows.find((r) => r.type === 'cover');
    if (cover && cover.model.subtitle) $('#report-name').value = cover.model.subtitle;
    else if (cover) $('#report-name').value = cover.model.title.replace(/^HSBC Sport\s*/i, '');
    $('#workspace').hidden = false;
    renderAll();
    loadThumbs();
    const hidden = state.deck.slides.filter((s) => s.hidden).length;
    setStatus(`${state.deck.slides.length} slides read${hidden ? `, ${hidden} hidden (not shown)` : ''}. Pick a layout for each slide you want, then download.`);
  } catch (e) {
    console.error(e);
    setStatus(`Sorry, that file couldn't be read: ${e.message}`, true);
  }
}

function makeRow(slide, type, include) {
  const model = extract(slide, type);
  const options = optionsFor(model);
  return { slide, type, model, options, choice: 0, look: DEFAULT_THEME, include: include && type !== 'cover' };
}

async function loadThumbs() {
  const pics = new Map();
  for (const r of state.rows) for (const e of r.slide.elements) if (e.kind === 'pic' && !pics.has(e.image)) pics.set(e.image, e);
  for (const [path, el] of pics) {
    const url = await thumbFor(state.deck, { image: path, crop: el.crop });
    if (url) {
      state.thumbs.set(path, url);
      document.querySelectorAll(`.pv-img-missing[data-img="${CSS.escape(path)}"]`).forEach((d) => {
        d.classList.remove('pv-img-missing');
        d.style.background = `url('${url}') center/cover no-repeat`;
      });
    }
  }
}

// ---------- rendering ----------
// Every layout is offered in every look. Each slide shows one scrolling strip per look;
// clicking any preview picks that layout in that look for that slide.
const LOOKS = Object.keys(THEMES);

function ctxFor(row, look) {
  let section = 0;
  for (const r of state.rows) {
    if (r.include && r.type === 'divider') section++;
    if (r === row) break;
  }
  return { section: Math.max(section, 1), label: `${STYLE.report}  |  ${$('#report-name').value.trim() || 'Report'}`, theme: THEMES[look] };
}

// strips are drawn when they scroll into view, so long decks stay quick
const lazy = new IntersectionObserver((entries) => {
  for (const e of entries) if (e.isIntersecting) { lazy.unobserve(e.target); e.target._draw(); }
}, { rootMargin: '600px 0px' });

function renderAll() {
  const list = $('#slides');
  list.innerHTML = '';
  state.rows.forEach((row, i) => {
    if (row.slide.hidden && !state.showHidden) return;
    list.appendChild(renderRow(row, i));
  });
  updateCount();
}

function optionButton(row, i, look, k) {
  const o = row.options[k];
  const chosen = row.look === look && row.choice === k;
  const b = document.createElement('button');
  b.type = 'button';
  b.className = 'option' + (chosen ? ' is-chosen' : '');
  b.setAttribute('aria-pressed', chosen);
  b.title = `${o.name}, ${THEMES[look].name}`;
  let preview;
  try { preview = renderPreview(renderWith(o, row.model, ctxFor(row, look)), state.thumbs); }
  catch (e) { console.error(e); preview = '<div class="pv-slide pv-error">Preview unavailable</div>'; }
  b.innerHTML = `<div class="pv-frame">${preview}</div><span class="option-name">${escapeHtml(o.name)}</span>`;
  b.addEventListener('click', () => {
    row.choice = k; row.look = look; row.include = true;
    const el = b.closest('.slide-row');
    el.classList.remove('is-off');
    el.querySelector('.include input').checked = true;
    el.querySelectorAll('.option.is-chosen').forEach((x) => { x.classList.remove('is-chosen'); x.setAttribute('aria-pressed', 'false'); });
    b.classList.add('is-chosen'); b.setAttribute('aria-pressed', 'true');
    el.querySelector('.chosen-label').textContent = `${o.name} · ${THEMES[look].name}`;
    updateCount();
  });
  return b;
}

function renderRow(row, i) {
  const el = document.createElement('section');
  el.className = 'slide-row' + (row.include ? '' : ' is-off');
  el.dataset.index = i;
  const typeOpts = Object.entries(TYPES).map(([k, v]) => `<option value="${k}" ${k === row.type ? 'selected' : ''}>${v}</option>`).join('');
  const heading = row.model.title || row.model.subtitle || '(no title)';
  const chosenName = row.options[row.choice] ? `${row.options[row.choice].name} · ${THEMES[row.look].name}` : '';
  el.innerHTML = `
    <header class="row-head">
      <label class="include"><input type="checkbox" ${row.include ? 'checked' : ''} aria-label="Include slide ${row.slide.number}"> <span>Slide ${row.slide.number}${row.slide.hidden ? ' (hidden)' : ''}</span></label>
      <p class="row-title" title="${escapeAttr(heading)}">${escapeHtml(heading)}</p>
      <span class="chosen-label">${escapeHtml(chosenName)}</span>
      <label class="type">Slide type <select>${typeOpts}</select></label>
    </header>
    <div class="looks"></div>`;
  const looks = el.querySelector('.looks');
  if (!row.options.length) looks.innerHTML = '<p class="muted">No layouts for this slide type.</p>';
  else for (const look of LOOKS) {
    const t = THEMES[look];
    const lr = document.createElement('div');
    lr.className = 'look-row';
    lr.innerHTML = `<p class="look-name"><span class="swatch" style="background:#${t.bg}"><i style="background:#${t.glow ? t.glow[0] : t.panel}"></i><i style="background:#${t.surface}"></i><i style="background:#${t.accent}"></i></span>${escapeHtml(t.name)}</p><div class="strip" role="group" aria-label="${escapeAttr(t.name)} layouts"></div>`;
    const strip = lr.querySelector('.strip');
    strip.style.minHeight = '160px';
    strip._draw = () => { row.options.forEach((_, k) => strip.appendChild(optionButton(row, i, look, k))); strip.style.minHeight = ''; };
    lazy.observe(strip);
    looks.appendChild(lr);
  }
  el.querySelector('.include input').addEventListener('change', (e) => { row.include = e.target.checked; el.classList.toggle('is-off', !row.include); updateCount(); });
  el.querySelector('.type select').addEventListener('change', (e) => {
    const look = row.look;
    state.rows[i] = makeRow(row.slide, e.target.value, true);
    state.rows[i].look = look;
    rerender(i);
  });
  return el;
}

function rerender(i) {
  const old = document.querySelector(`.slide-row[data-index="${i}"]`);
  if (old) old.replaceWith(renderRow(state.rows[i], i));
  updateCount();
}
function updateCount() {
  const n = state.rows.filter((r) => r.include && r.options.length).length;
  $('#download').disabled = n === 0;
  $('#download').textContent = n ? `Download ${n} refreshed slide${n === 1 ? '' : 's'}` : 'Choose some slides first';
}

// ---------- download ----------
async function download() {
  const chosen = state.rows.filter((r) => r.include && r.options.length);
  if (!chosen.length) return;
  const btn = $('#download');
  btn.disabled = true;
  setStatus(`Building ${chosen.length} slides…`);
  try {
    const entries = chosen.map((r) => ({ background: THEMES[r.look].bg, boxes: renderWith(r.options[r.choice], r.model, ctxFor(r, r.look)) }));
    const pres = await buildPptx(entries, { PptxGenJS: window.PptxGenJS, getImageData: imageMaker(state.deck), getGradient: makeGradient });
    const bytes = await finishFile(await pres.write({ outputType: 'arraybuffer' }), window.JSZip);
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${state.fileName} - refreshed.pptx`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 60000);
    setStatus(`Done. "${a.download}" has been downloaded. Open it in PowerPoint and paste the slides into your deck.`);
  } catch (e) {
    console.error(e);
    setStatus(`Something went wrong building the file: ${e.message}`, true);
  } finally { updateCount(); }
}

// previews have a fixed width, so the scale is fixed too (see .pv-frame in css/app.css)
document.documentElement.style.setProperty('--pv-scale', String(220 / 1333));

// ---------- helpers & wiring ----------
function setStatus(msg, isError = false) { const s = $('#status'); s.textContent = msg; s.classList.toggle('is-error', isError); }
function escapeHtml(s) { return String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c])); }
function escapeAttr(s) { return escapeHtml(s).replace(/"/g, '&quot;'); }

const drop = $('#drop');
$('#file').addEventListener('change', (e) => e.target.files[0] && load(e.target.files[0]));
['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('is-over'); }));
['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('is-over'); }));
drop.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) load(f); });
$('#download').addEventListener('click', download);
$('#show-hidden').addEventListener('change', (e) => { state.showHidden = e.target.checked; renderAll(); });
$('#report-name').addEventListener('change', renderAll);
$('#select-all').addEventListener('click', () => { state.rows.forEach((r) => { if (!r.slide.hidden && r.type !== 'cover') r.include = true; }); renderAll(); });
$('#select-none').addEventListener('click', () => { state.rows.forEach((r) => { r.include = false; }); renderAll(); });

// test hook: ?demo=path/to/deck.pptx loads a deck on open (used for local checks)
const demo = new URLSearchParams(location.search).get('demo');
if (demo) fetch(demo).then((r) => r.blob()).then((b) => load(new File([b], demo.split('/').pop())));
