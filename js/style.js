// House style for every slide the tool builds.
// Change a value here and every template picks it up.

export const STYLE = {
  slide: { w: 13.333, h: 7.5 },            // inches, 16:9 widescreen
  margin: { x: 0.43, top: 0.43 },
  fonts: {
    body: 'Univers Next for HSBC Light',   // main text
    bold: 'Univers Next for HSBC',         // headline figures
  },
  colors: {
    red: 'C00000',        // HSBC red: stat cards, highlights, bullets
    dark: '333333',       // text and card top bars
    text: '333333',
    soft: '4D4D4D',       // body copy on cards
    grey: '767676',       // labels and captions
    mid: 'A6A6A6',
    light: 'D9D9D9',      // dividers and hairlines
    card: 'F3F3F3',       // card background
    white: 'FFFFFF',
    gold: 'B8952E',       // Blockbuster
    good: '3E701A',       // comparison going up
    bad: 'C00000',        // comparison going down
    panelLabel: 'F2C4C4', // small text on red panels
    panelRule: 'D9666B',  // thin rules on red panels
  },
  // series colours for rebuilt charts (current year first)
  chartSeries: ['C00000', 'A6A6A6', '333333', 'D9D9D9', 'B8952E'],
  // content framework pillars
  pillars: {
    'blockbuster': 'B8952E',
    'partnership storytelling': '333333',
    'insider pulse': 'A6A6A6',
    'outside framework': 'C00000',
  },
  report: 'HSBC Sport', // shown on dividers, e.g. "HSBC Sport  |  Q3 Report 2026"
  maxImagePx: 2200,      // photos are re-saved no wider than this
};
