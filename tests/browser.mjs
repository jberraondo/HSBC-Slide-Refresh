// End-to-end check in Chrome:  node tests/browser.mjs path/to/deck.pptx [--theme=svns]
// Serves the site, loads the deck through the page's ?demo= hook, saves a screenshot,
// clicks Download and checks a .pptx arrives. Needs Google Chrome installed.
import puppeteer from 'puppeteer-core';
import { createServer } from 'http';
import { readFileSync, existsSync, mkdirSync, readdirSync, statSync, rmSync } from 'fs';
import { join, extname, dirname, relative } from 'path';
import { fileURLToPath } from 'url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const deck = (process.argv[2] && !process.argv[2].startsWith('--')) ? process.argv[2] : 'tests/sample.pptx';
const theme = (process.argv.find((a) => a.startsWith('--theme=')) || '').split('=')[1];
const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.pptx': 'application/octet-stream' };
const server = createServer((req, res) => {
  const p = join(root, decodeURIComponent(req.url.split('?')[0]));
  if (!p.startsWith(root) || !existsSync(p) || statSync(p).isDirectory()) {
    const idx = join(root, 'index.html');
    if (req.url.split('?')[0] === '/') { res.writeHead(200, { 'content-type': 'text/html' }); return res.end(readFileSync(idx)); }
    res.writeHead(404); return res.end();
  }
  res.writeHead(200, { 'content-type': types[extname(p)] || 'application/octet-stream' }); res.end(readFileSync(p));
}).listen(8765);

const out = join(root, 'tests/out/browser'); rmSync(out, { recursive: true, force: true }); mkdirSync(out, { recursive: true });
const browser = await puppeteer.launch({ executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', headless: 'new', args: ['--no-first-run'] });
const page = await browser.newPage();
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
await page.setViewport({ width: 1366, height: 900 });
const cdp = await page.target().createCDPSession();
await cdp.send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath: out });
await page.goto(`http://localhost:8765/?demo=/${relative(root, deck)}`, { waitUntil: 'networkidle0' });
await page.waitForSelector('.slide-row', { timeout: 60000 });
await new Promise((r) => setTimeout(r, 2500));
const rows = await page.$$eval('.slide-row', (els) => els.length);
const opts = await page.$$eval('.option', (els) => els.length);
const looks = await page.$$eval('.look-row', (els) => els.length);
// draw every strip, then pick the first SVNS Dark layout on the 4th slide
for (let y = 0; y < 60000; y += 800) { await page.evaluate((yy) => window.scrollTo(0, yy), y); await new Promise((r) => setTimeout(r, 60)); }
const optionsDrawn = await page.$$eval('.option', (els) => els.length);
const picked = await page.evaluate(() => {
  const row = document.querySelectorAll('.slide-row')[3];
  const strip = [...row.querySelectorAll('.look-row')].find((r) => r.textContent.includes('SVNS'));
  strip.querySelector('.option').click();
  return row.querySelector('.chosen-label').textContent;
});
await page.screenshot({ path: join(out, 'page-top.png') });
await page.evaluate(() => window.scrollTo(0, 900));
await page.screenshot({ path: join(out, 'page-mid.png') });
await page.click('#download');
let file = null;
for (let i = 0; i < 120 && !file; i++) { await new Promise((r) => setTimeout(r, 1000)); file = readdirSync(out).find((f) => f.endsWith('.pptx')); }
console.log(JSON.stringify({ rows, looks, optionsDrawn, picked, download: file || null, size: file ? statSync(join(out, file)).size : 0, status: await page.$eval('#status', (e) => e.textContent), errors }, null, 1));
await browser.close(); server.close();
