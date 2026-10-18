const { C, DEPT } = require('./theme.js');

const FOOT = 'Corporate Technology Services  |  Seamless AV';
// Geometry copied from the September deck so every slide lands in the same place.
const G = {
  eyebrow: { x: 0.55, y: 0.40, w: 7.00, h: 0.30 },
  title:   { x: 0.55, y: 0.70, w: 8.00, h: 0.75 },
  body:    { x: 0.55, y: 1.60 },
  panel:   { x: 6.85, y: 1.60, w: 2.60, h: 3.45 },
  panelLabel: { x: 7.05, y: 1.78, w: 2.20, h: 0.25 },
  panelText:  { x: 7.05, y: 2.08, w: 2.20, h: 2.85 },
  wide:    { x: 0.55, w: 8.90 },
};

const money = (v, dp = 0) =>
  (v < 0 ? '(' : '') + Math.abs(v).toLocaleString('en-AU', { minimumFractionDigits: dp, maximumFractionDigits: dp }) + (v < 0 ? ')' : '');
const k = (v) => money(v / 1000);
const pct = (v, dp = 1) => (v * 100).toFixed(dp) + '%';
const signed = (v) => (v >= 0 ? '+' : '-') + Math.abs(v).toLocaleString('en-AU', { maximumFractionDigits: 0 });
const varColour = (v) => (v >= 0 ? C.good : C.bad);

/** Eyebrow + title + footer, the frame every content slide shares. */
function head(slide, eyebrow, title) {
  slide.addText(eyebrow.toUpperCase(), {
    ...G.eyebrow, isTextBox: true, margin: 0, fontSize: 9.5, bold: true,
    color: C.eyebrow, charSpacing: 2.2, valign: 'middle',
  });
  slide.addText(title, {
    ...G.title, isTextBox: true, margin: 0, fontSize: 27, bold: true,
    color: C.navy, valign: 'middle',
  });
}

// Roughly 34 characters fit on a line in the panel at 10pt. Step down as the
// copy grows so the last bullet always lands inside the tint.
const sizeFor = (b) => { const n = b.join(' ').length; return n > 560 ? 8.5 : n > 440 ? 9 : n > 340 ? 9.5 : 10; };
const lineFor = (b) => { const n = b.join(' ').length; return n > 560 ? 11 : n > 440 ? 11.5 : 13; };

/** The pale commentary panel on the right. Bullets are short conclusions. */
function commentary(slide, bullets, opts = {}) {
  const box = { ...G.panel, ...(opts.panel || {}) };
  slide.addShape('rect', { ...box, fill: { color: C.mint }, line: { type: 'none' }, objectName: 'commentary panel' });
  slide.addText('COMMENTARY', {
    ...G.panelLabel, y: box.y + 0.18, isTextBox: true, margin: 0,
    fontSize: 9, bold: true, color: C.teal, charSpacing: 1.8,
  });
  slide.addText(
    bullets.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < bullets.length - 1 } })),
    {
      ...G.panelText, y: box.y + 0.48, h: box.h - 0.66, isTextBox: true, margin: 0,
      fontSize: opts.fontSize || sizeFor(bullets), color: C.body,
      lineSpacing: opts.fontSize ? 13 : lineFor(bullets),
      paraSpaceAfter: 5, valign: 'top', shrinkText: true,
    });
}

/** A caption under a table or chart. */
function note(slide, text, o) {
  slide.addText(text, {
    x: o.x, y: o.y, w: o.w, h: o.h || 0.4, isTextBox: true, margin: 0,
    fontSize: 8.5, italic: true, color: C.muted, valign: 'top',
  });
}

/** Big-number tile, as used on the risk and receivables slides. */
function tile(slide, o) {
  slide.addShape('rect', { x: o.x, y: o.y, w: o.w, h: o.h, fill: { color: C.tile }, line: { type: 'none' } });
  slide.addText(o.value, {
    x: o.x + 0.15, y: o.y + 0.12, w: o.w - 0.3, h: 0.42, isTextBox: true, margin: 0,
    fontSize: 21, bold: true, color: o.colour || C.navy, valign: 'middle',
  });
  slide.addText(o.label, {
    x: o.x + 0.15, y: o.y + 0.55, w: o.w - 0.3, h: o.h - 0.66, isTextBox: true, margin: 0,
    fontSize: 8.5, color: C.muted, valign: 'top',
  });
}

/** Table styling: navy header, hairline rules, bold total row. */
const HEAD = { fill: C.navy, color: C.white, bold: true, fontSize: 9 };
const CELL = { fontSize: 9, color: C.body };
function row(cells, o = {}) {
  return cells.map((c) => {
    const t = typeof c === 'object' && c !== null ? c : { text: String(c) };
    return {
      text: t.text,
      options: {
        ...CELL, align: t.align || 'right', bold: o.bold || t.bold || false,
        color: t.color || (o.bold ? C.navy : C.body),
        fill: o.fill ? { color: o.fill } : undefined,
        valign: 'middle', margin: [2, 5, 2, 5],
      },
    };
  });
}
function headRow(cells, align) {
  return cells.map((c, i) => ({
    text: typeof c === 'object' ? c.text : c,
    options: { ...HEAD, align: align || (i === 0 ? 'left' : 'right'), valign: 'middle', margin: [3, 5, 3, 5] },
  }));
}

/** Chart defaults, so every chart in the deck reads the same way. */
const CHART = {
  showLegend: false, showTitle: false,
  catAxisLabelColor: C.body, catAxisLabelFontSize: 8.5, catAxisLabelFontFace: '+mn-lt',
  valAxisLabelColor: C.muted, valAxisLabelFontSize: 8.5, valAxisLabelFontFace: '+mn-lt',
  dataLabelColor: C.body, dataLabelFontSize: 8.5, dataLabelFontFace: '+mn-lt',
  catAxisLineShow: false, valAxisLineShow: false,
  valGridLine: { color: C.rule, size: 0.5 },
  catGridLine: { style: 'none' },
  chartColors: [C.teal],
};

module.exports = { G, FOOT, money, k, pct, signed, varColour, head, commentary, note, tile, row, headRow, CHART, HEAD, CELL, DEPT, C };
