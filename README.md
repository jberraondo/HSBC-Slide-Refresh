# HSBC Slide Refresh

A free web tool that turns PowerPoint report slides into the HSBC Sport report look: red stat cards, grey content cards with a dark top bar, native charts in house colours, and red-panel section dividers.

**Use it:** https://jberraondo.github.io/HSBC-Slide-Refresh/

## How it works

1. Open the page and add a `.pptx` file.
2. The tool reads each slide (text, numbers, photos, chart data), works out what kind of slide it is, and shows two or three layouts for it.
3. Tick the slides you want, pick a layout for each, and download a new `.pptx`.
4. Open it in PowerPoint and paste the slides into your deck.

Everything runs in your browser. **Your file is never uploaded anywhere**, so it's safe for client decks.

## Themes

Pick a theme at the top of the slide list and every layout and preview switches to it:

| Theme | Look |
|---|---|
| HSBC Classic | White slides, red stat cards, grey cards (the Q3 report look) |
| SVNS Dark (red glow) | Near-black slides with a red glow, as in the SVNS / Dubai decks |
| Midnight | Pure black with bold red figures |
| Charcoal | Dark grey with red panels |
| All Red | Red slides with white cards |
| Soft Grey | Light grey slides, white cards, charcoal dividers |
| Mono | Black and white, red only for drops |
| Blockbuster Gold | Black with gold highlights |

The chosen theme is remembered on that computer. To add or change a theme, edit `THEMES` in [`js/style.js`](js/style.js).

## What it recognises

| Slide type | Layouts |
|---|---|
| Cover | Panel and photo, full photo fading to title, glow, title left / photo right |
| Contents | Panel list with photo, photo and numbered grid, numbered list |
| Section divider | Panel with number, glow with big title, photo fading into title, full photo with band, clean title with photo strip |
| Headline figures (e.g. "184M Impressions, +128% on Q3 2025") | Stat cards, one big number, clean row, figure list, colour block |
| Content cards (photo + title + "Key Results") | Card grid (1 to 8 cards), photo cards with text on top, image and text rows, wide hero cards, single feature |
| Charts (line, bar, doughnut) | Stat cards above each chart, charts on cards, big chart with insight panel, hero chart, share-of-total bars, donut on a card, share cards |
| Learnings / text | Numbered cards, insight statement, numbered circles, side by side, headings and bullets, copy over photo fade |
| Tables | Styled tables with copy |

If a slide is put in the wrong type, change it with the **Slide type** menu and the layouts update.

## Good to know

- It builds **new slides** from what's already in your slides. Expect about 80% of the finish, and check figures and wording before sending.
- Text is set in **Univers Next for HSBC**, so the file looks right on laptops with the HSBC fonts installed. The previews on the page use a stand-in font.
- Charts are **native PowerPoint charts**: right-click, Edit Data. They show in PowerPoint; Keynote may show them blank.
- Photos are cropped to their new frames and saved as normal-size JPEGs, which avoids "The picture can't be displayed" when pasting between decks.
- Grouped shapes, SmartArt and EMF/WMF images are only partly carried across.

## Changing the house style

All colours, fonts and sizes are in [`js/style.js`](js/style.js). Change a value there and every layout picks it up. The layouts themselves are in [`js/templates.js`](js/templates.js): each option is a small function that returns boxes (rectangles, text, images, charts) in inches on a 13.33 x 7.5 slide.

## Files

| File | What it does |
|---|---|
| `index.html`, `css/app.css`, `js/app.js` | The page |
| `js/parse.js` | Reads the `.pptx` (slides, text, pictures, chart data, tables) |
| `js/classify.js` | Works out each slide's type and pulls out its pieces |
| `js/templates.js` | The layout options |
| `js/build.js` | Writes the new `.pptx` with [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) |
| `js/preview.js` | Draws the small previews |
| `js/images.js` | Crops and resizes photos in the browser |
| `js/style.js` | House style |

Libraries are loaded from public CDNs at fixed versions: JSZip 3.10.1 and PptxGenJS 3.12.0.

## Testing (for whoever maintains it)

```bash
npm install                       # test-only libraries
python3 tests/make_sample.py      # rebuilds tests/sample.pptx (made-up content, needs python-pptx + pillow)
node tests/run.mjs                # every layout for every slide in the sample -> tests/out/sample-refreshed.pptx
node tests/run.mjs my-deck.pptx   # same for any deck
node tests/browser.mjs            # end-to-end in Chrome: load, preview, download
```

Put real client decks in `tests/private/` (ignored by git) so they're never committed.

## Publishing

The site is plain static files served by GitHub Pages (Settings → Pages → Deploy from branch → `main`, root). Any push to `main` updates the live page within a minute or two.

