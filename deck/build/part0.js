const B = require('./build.js');
const L = require('./lib.js');
const { C, DEPT } = require('./theme.js');
const { head, commentary, note, tile, row, headRow, money, k, pct, signed, varColour } = L;
const { pres, D, P, rev, gp, oh, eb, np, depts, dTot, byRev, ytdMargin, breakEven, revRun, MONTH, SHORT } = B;

/** White on a dark fill, navy on a light one, so every card stays readable. */
function inkOn(hex) {
  const n = parseInt(hex, 16);
  const lum = (0.299 * ((n >> 16) & 255) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255;
  return lum > 0.6 ? C.navy : C.white;
}

/* ============================ introduction: the business at a glance ============ */
function introSlide() {
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Introduction', `Where the business stands, ${SHORT}`);

  // One card per department, each in that department's own colour.
  const order = ['Support', 'Events', 'Video', 'Integration', 'Consulting'];
  order.forEach((name, i) => {
    const d = depts.find((x) => x.name === name);
    const fill = DEPT[name];
    const ink = inkOn(fill);
    const soft = ink === C.white ? C.lt2 || 'E3E6EA' : C.navy;
    const x = 0.55 + i * 1.79;
    s.addShape('rect', { x, y: 1.62, w: 1.62, h: 1.78, fill: { color: fill }, line: { type: 'none' }, objectName: `${name} card` });
    s.addText(name.toUpperCase(), { x: x + 0.12, y: 1.74, w: 1.38, h: 0.22, isTextBox: true, margin: 0, fontSize: 8.5, bold: true, color: ink, charSpacing: 1.2 });
    s.addText('$' + k(d.revenue) + 'k', { x: x + 0.12, y: 1.98, w: 1.38, h: 0.40, isTextBox: true, margin: 0, fontSize: 19, bold: true, color: ink, valign: 'middle' });
    s.addText('revenue', { x: x + 0.12, y: 2.36, w: 1.38, h: 0.20, isTextBox: true, margin: 0, fontSize: 8, color: soft, transparency: 15 });
    s.addText(pct(d.gm_pct, 0) + ' gross margin', { x: x + 0.12, y: 2.62, w: 1.38, h: 0.20, isTextBox: true, margin: 0, fontSize: 9, color: ink });
    s.addText(`${d.result >= 0 ? '+' : ''}$${k(d.result)}k after overhead`, { x: x + 0.12, y: 2.86, w: 1.38, h: 0.22, isTextBox: true, margin: 0, fontSize: 9, bold: true, color: ink });
    s.addText(`${d.revenue >= d.budget ? 'Ahead of' : 'Behind'} budget by $${k(Math.abs(d.revenue - d.budget))}k`,
      { x: x + 0.12, y: 3.08, w: 1.38, h: 0.24, isTextBox: true, margin: 0, fontSize: 7.5, color: soft });
  });

  // Company line underneath, in the house navy.
  s.addShape('rect', { x: 0.55, y: 3.58, w: 8.90, h: 0.62, fill: { color: C.tile }, line: { type: 'none' }, objectName: 'company strip' });
  const stats = [
    ['Revenue', '$' + k(rev.month) + 'k', `budget $${k(rev.month_budget)}k`],
    ['Gross margin', pct(gp.month / rev.month), `last Sep ${pct(gp.month_ly / rev.month_ly)}`],
    ['Overheads', '$' + k(-oh.month) + 'k', `budget $${k(-oh.month_budget)}k`],
    ['EBITDA', '$' + k(eb.month) + 'k', `budget $${k(eb.month_budget)}k`],
    ['Break even needs', '$' + k(breakEven) + 'k', `running at $${k(revRun)}k`],
  ];
  stats.forEach(([lab, val, sub], i) => {
    const x = 0.70 + i * 1.76;
    s.addText(lab, { x, y: 3.64, w: 1.6, h: 0.18, isTextBox: true, margin: 0, fontSize: 7.5, color: C.muted, charSpacing: 0.6 });
    s.addText(val, { x, y: 3.80, w: 1.6, h: 0.24, isTextBox: true, margin: 0, fontSize: 13, bold: true, color: i === 3 ? C.bad : C.navy, valign: 'middle' });
    s.addText(sub, { x, y: 4.02, w: 1.6, h: 0.16, isTextBox: true, margin: 0, fontSize: 7, color: C.muted });
  });

  s.addText([
    { text: 'The month in one line. ', options: { bold: true, color: C.navy } },
    { text: `Revenue beat budget by $${k(rev.month - rev.month_budget)}k and the company still lost $${k(-eb.month)}k of EBITDA. Support is the only department that covers its own overhead. Integration delivered the revenue beat at ${pct(depts[3].gm_pct)} gross margin, so it added turnover and almost no profit. At the ${pct(ytdMargin)} margin we actually deliver, break even needs $${k(breakEven)}k of revenue a month and we are running at $${k(revRun)}k.` },
  ], { x: 0.55, y: 4.36, w: 8.90, h: 0.70, isTextBox: true, margin: 0, fontSize: 10, color: C.body, lineSpacing: 14, valign: 'top' });

  s.addNotes(`The introduction slide, colour coded by department: Support in red, Events and Video in the Production teals, Integration and Consulting in the Consulting blues.\n\n` +
    order.map((nm) => { const d = depts.find((x) => x.name === nm); return `${nm}: revenue ${money(d.revenue)}, gross margin ${pct(d.gm_pct)}, result ${money(d.result)} after all overhead, ${d.revenue >= d.budget ? 'ahead of' : 'behind'} budget by ${money(Math.abs(d.revenue - d.budget))}.`; }).join('\n') +
    `\n\nCompany revenue ${money(rev.month)} against ${money(rev.month_budget)} budget. Gross margin ${pct(gp.month / rev.month)}. Overheads ${money(-oh.month)}. EBITDA ${money(eb.month)}.\n\nIf the room only hears one thing from this pack, it is the last figure on the strip. At our real margin we need ${money(breakEven)} of revenue every month to break even and we are averaging ${money(revRun)}.`);
  return s;
}

/* ============================ slide 8 rebuilt as drawn bars ==================== */
/** Charts put category labels next to the zero line when values go negative, which
 *  sits them on top of the bars. Drawing the bars keeps the labels clear. */
function barBlock(s, o) {
  const { x, y, w, rows, max, min, fmt, rowH = 0.30, labelW = 1.05 } = o;
  const plotX = x + labelW, plotW = w - labelW - 0.55;
  const span = max - min;
  const zero = plotX + ((0 - min) / span) * plotW;
  if (min < 0) s.addShape('line', { x: zero, y, w: 0, h: rows.length * rowH, line: { color: C.rule, width: 0.75 } });
  rows.forEach((r, i) => {
    const ry = y + i * rowH;
    const len = (Math.abs(r.value) / span) * plotW;
    const bx = r.value >= 0 ? zero : zero - len;
    s.addText(r.name, { x, y: ry + 0.02, w: labelW - 0.08, h: rowH - 0.04, isTextBox: true, margin: 0, fontSize: 9, color: C.body, align: 'right', valign: 'middle' });
    s.addShape('rect', { x: bx, y: ry + 0.055, w: Math.max(len, 0.012), h: rowH - 0.11, fill: { color: r.colour }, line: { type: 'none' }, objectName: `${r.name} bar` });
    s.addText(fmt(r.value), {
      x: r.value >= 0 ? bx + len + 0.04 : bx - 0.60, y: ry + 0.02, w: 0.56, h: rowH - 0.04,
      isTextBox: true, margin: 0, fontSize: 9, bold: true, color: C.navy,
      align: r.value >= 0 ? 'left' : 'right', valign: 'middle',
    });
  });
}

// Full year budget figures, needed for the EBITDA risk on the cost and profit slide.
const EB = {
  q1_actual: B.eb.ytd,
  q1_budget: B.eb.ytd_budget,
  fy_budget: 397251,          // Budget FY27 row 176, twelve months
  fy26_actual: 235273,        // Actual FY26 row 176, twelve months
  fy26_q1: -85404,            // Actual FY26 row 176, Jul to Sep 2025
};
EB.shortfall = EB.q1_actual - EB.q1_budget;
EB.still_needed = EB.fy_budget - EB.q1_actual;
EB.budget_assumed_rest = EB.fy_budget - EB.q1_budget;
EB.per_month_needed = EB.still_needed / 9;
EB.per_month_assumed = EB.budget_assumed_rest / 9;

module.exports = { introSlide, barBlock, inkOn, EB };
