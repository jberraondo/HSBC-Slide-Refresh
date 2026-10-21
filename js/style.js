// House style. Fonts, sizes and the colour themes every layout uses.
// To add a theme, copy one below, give it a new key and name, and change the colours.

export const STYLE = {
  slide: { w: 13.333, h: 7.5 },            // inches, 16:9 widescreen
  margin: { x: 0.43, top: 0.43 },
  fonts: {
    body: 'Univers Next for HSBC Light',   // all text
    bold: 'Univers Next for HSBC',
  },
  report: 'HSBC Sport', // shown on dividers, e.g. "HSBC Sport  |  Q3 Report 2026"
  maxImagePx: 2200,     // photos are re-saved no wider than this
};

// Colour roles each theme fills in:
//   bg / text / muted / soft        slide background, main text, labels, body copy
//   accent / onAccent               highlight colour (bullets, rules, labels) and text placed on it
//   surface / onSurface / bar       cards, text on cards, the thin bar on top of cards
//   cardAccent                      highlight used on cards (defaults to accent)
//   statFill / statNum / onStat     headline-figure cards: fill, the big number, the other text
//   panel / onPanel / panelMuted    dividers and covers; glow = [glow colour, base colour] or null
//   line / good / bad               hairlines, comparisons going up / down
//   chart / chartText / chartGrid   series colours (first is the highlight), axis text, gridlines
//   fade                            colour photos fade into on "photo fade" layouts
//   pillars                         content framework colours
const pillars = (bb, ps, ip, of) => ({ 'blockbuster': bb, 'partnership storytelling': ps, 'insider pulse': ip, 'outside framework': of });

export const THEMES = {
  classic: {
    name: 'HSBC Classic', dark: false,
    bg: 'FFFFFF', text: '333333', muted: '767676', soft: '4D4D4D',
    accent: 'C00000', onAccent: 'FFFFFF',
    surface: 'F3F3F3', onSurface: '333333', bar: '333333',
    statFill: 'C00000', statNum: 'FFFFFF', onStat: 'FFFFFF',
    panel: 'C00000', onPanel: 'FFFFFF', panelMuted: 'F2C4C4', panelRule: 'D9666B', glow: ['E0303A', '9A0000'],
    line: 'D9D9D9', good: '3E701A', bad: 'C00000',
    chart: ['C00000', 'A6A6A6', '333333', 'D9D9D9', 'B8952E'], chartText: '767676', chartGrid: 'D9D9D9',
    fade: 'FFFFFF', pillars: pillars('B8952E', '333333', 'A6A6A6', 'C00000'),
  },
  svns: {
    name: 'SVNS Dark (red glow)', dark: true,
    bg: '0B0B0B', text: 'FFFFFF', muted: 'A6A6A6', soft: 'D9D9D9',
    accent: 'E2231A', onAccent: 'FFFFFF',
    surface: '1C1C1C', onSurface: 'FFFFFF', bar: 'E2231A',
    statFill: '1C1C1C', statNum: 'FFFFFF', onStat: 'D9D9D9',
    panel: '0B0B0B', onPanel: 'FFFFFF', panelMuted: 'BFBFBF', panelRule: '4A1515', glow: ['8E1212', '0B0B0B'],
    line: '333333', good: '7BC86C', bad: 'FF6B6B',
    chart: ['E2231A', '8C8C8C', 'FFFFFF', '4D4D4D', 'D4AF37'], chartText: 'BFBFBF', chartGrid: '333333',
    fade: '0B0B0B', pillars: pillars('D4AF37', 'FFFFFF', '8C8C8C', 'E2231A'),
  },
  midnight: {
    name: 'Midnight (black)', dark: true,
    bg: '000000', text: 'FFFFFF', muted: '9E9E9E', soft: 'D0D0D0',
    accent: 'DB0011', onAccent: 'FFFFFF',
    surface: '141414', onSurface: 'FFFFFF', bar: 'FFFFFF',
    statFill: 'DB0011', statNum: 'FFFFFF', onStat: 'FFFFFF',
    panel: '000000', onPanel: 'FFFFFF', panelMuted: 'B3B3B3', panelRule: '333333', glow: null,
    line: '2E2E2E', good: '7BC86C', bad: 'FF6B6B',
    chart: ['DB0011', '7F7F7F', 'FFFFFF', '3D3D3D', 'D4AF37'], chartText: 'B3B3B3', chartGrid: '2E2E2E',
    fade: '000000', pillars: pillars('D4AF37', 'FFFFFF', '7F7F7F', 'DB0011'),
  },
  charcoal: {
    name: 'Charcoal', dark: true,
    bg: '2B2B2B', text: 'FFFFFF', muted: 'B3B3B3', soft: 'E0E0E0',
    accent: 'E2231A', onAccent: 'FFFFFF',
    surface: '3A3A3A', onSurface: 'FFFFFF', bar: 'E2231A',
    statFill: 'C00000', statNum: 'FFFFFF', onStat: 'FFFFFF',
    panel: 'C00000', onPanel: 'FFFFFF', panelMuted: 'F2C4C4', panelRule: 'D9666B', glow: null,
    line: '4A4A4A', good: '8FD07F', bad: 'FF7A7A',
    chart: ['E2231A', 'A6A6A6', 'FFFFFF', '5C5C5C', 'D4AF37'], chartText: 'C8C8C8', chartGrid: '4A4A4A',
    fade: '2B2B2B', pillars: pillars('D4AF37', 'FFFFFF', 'A6A6A6', 'E2231A'),
  },
  allred: {
    name: 'All Red', dark: true,
    bg: 'C00000', text: 'FFFFFF', muted: 'F2C4C4', soft: 'FBE3E5',
    accent: 'FFFFFF', onAccent: 'C00000', cardAccent: 'C00000',
    surface: 'FFFFFF', onSurface: '333333', bar: '333333',
    statFill: 'FFFFFF', statNum: 'C00000', onStat: '333333',
    panel: 'C00000', onPanel: 'FFFFFF', panelMuted: 'F2C4C4', panelRule: 'D9666B', glow: ['E0303A', '8E0000'],
    line: 'D9666B', good: '3E701A', bad: 'C00000',
    chart: ['C00000', 'A6A6A6', '333333', 'D9D9D9', 'B8952E'], chartText: '767676', chartGrid: 'E6E6E6',
    fade: 'C00000', pillars: pillars('B8952E', '333333', 'A6A6A6', 'C00000'),
  },
  softgrey: {
    name: 'Soft Grey', dark: false,
    bg: 'F3F3F3', text: '333333', muted: '767676', soft: '4D4D4D',
    accent: 'C00000', onAccent: 'FFFFFF',
    surface: 'FFFFFF', onSurface: '333333', bar: 'C00000',
    statFill: 'FFFFFF', statNum: 'C00000', onStat: '333333',
    panel: '333333', onPanel: 'FFFFFF', panelMuted: 'BFBFBF', panelRule: '555555', glow: null,
    line: 'D9D9D9', good: '3E701A', bad: 'C00000',
    chart: ['C00000', 'A6A6A6', '333333', 'D9D9D9', 'B8952E'], chartText: '767676', chartGrid: 'E3E3E3',
    fade: 'F3F3F3', pillars: pillars('B8952E', '333333', 'A6A6A6', 'C00000'),
  },
  mono: {
    name: 'Mono', dark: false,
    bg: 'FFFFFF', text: '1A1A1A', muted: '767676', soft: '4D4D4D',
    accent: '1A1A1A', onAccent: 'FFFFFF',
    surface: 'F3F3F3', onSurface: '1A1A1A', bar: '1A1A1A',
    statFill: '1A1A1A', statNum: 'FFFFFF', onStat: 'D9D9D9',
    panel: '1A1A1A', onPanel: 'FFFFFF', panelMuted: 'BFBFBF', panelRule: '4D4D4D', glow: null,
    line: 'D9D9D9', good: '1A1A1A', bad: 'C00000',
    chart: ['1A1A1A', 'A6A6A6', 'D9D9D9', '595959', 'C00000'], chartText: '767676', chartGrid: 'E3E3E3',
    fade: 'FFFFFF', pillars: pillars('C00000', '1A1A1A', 'A6A6A6', '595959'),
  },
  gold: {
    name: 'Blockbuster Gold', dark: true,
    bg: '0B0B0B', text: 'FFFFFF', muted: 'A6A6A6', soft: 'D9D9D9',
    accent: 'C9A44A', onAccent: '0B0B0B',
    surface: '1C1C1C', onSurface: 'FFFFFF', bar: 'C9A44A',
    statFill: '1C1C1C', statNum: 'C9A44A', onStat: 'D9D9D9',
    panel: '0B0B0B', onPanel: 'FFFFFF', panelMuted: 'C9A44A', panelRule: '4A3C17', glow: ['5C4A17', '0B0B0B'],
    line: '333333', good: '7BC86C', bad: 'FF6B6B',
    chart: ['C9A44A', '8C8C8C', 'FFFFFF', '4D4D4D', 'E2231A'], chartText: 'BFBFBF', chartGrid: '333333',
    fade: '0B0B0B', pillars: pillars('C9A44A', 'FFFFFF', '8C8C8C', 'E2231A'),
  },
};

export const DEFAULT_THEME = 'classic';
