const pptxgen = require('pptxgenjs');
const { THEME, C, DEPT } = require('./theme.js');
const L = require('./lib.js');
const { G, FOOT, money, k, pct, signed, varColour, head, commentary, note, tile, row, headRow, CHART } = L;
const D = require('./data.json');

const MONTH = D.setup_month;                 // "September 2026"
const SHORT = 'Sep 26';
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';                 // 10" x 5.625", same canvas as the September deck
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = 'Danica Nelson';
pres.title = `CTS Monthly Business Review - ${MONTH}`;
pres.subject = 'Monthly Business Review';

/* ------------------------------------------------------------------ layouts */
pres.defineSlideMaster({
  title: 'CTS_DARK', background: { color: C.navy },
  objects: [{ text: { text: FOOT, options: { x: 0.55, y: 5.18, w: 4.5, h: 0.25, fontSize: 8, color: C.grey, isTextBox: true, margin: 0, valign: 'middle' } } }],
});
pres.defineSlideMaster({
  title: 'CTS_LIGHT', background: { color: C.white },
  objects: [{ text: { text: FOOT, options: { x: 0.55, y: 5.18, w: 4.5, h: 0.25, fontSize: 8, color: C.muted, isTextBox: true, margin: 0, valign: 'middle' } } }],
  slideNumber: { x: 8.85, y: 5.18, w: 0.60, h: 0.25, fontSize: 8, color: C.muted, align: 'right' },
});

/* ------------------------------------------------------------- derived facts */
const P = D.pl;
const rev = P.revenue, gp = P.gross_profit, oh = P.overheads, eb = P.ebitda, np = P.net_profit;
const ytdMargin = gp.ytd / rev.ytd;
const ohRun = Math.abs(oh.ytd) / 3;                       // overhead a month, year to date
const breakEven = ohRun / ytdMargin;                      // revenue needed to cover it
const revRun = rev.ytd / 3;
const depts = D.dept;
const dTot = depts.reduce((a, x) => ({
  revenue: a.revenue + x.revenue, budget: a.budget + x.budget, gross_profit: a.gross_profit + x.gross_profit,
  all_overhead: a.all_overhead + x.all_overhead, result: a.result + x.result,
}), { revenue: 0, budget: 0, gross_profit: 0, all_overhead: 0, result: 0 });
const byRev = [...depts].sort((a, b) => b.revenue - a.revenue);
const util = D.util;
const utilTot = util.reduce((a, u) => ({ c: a.c + u.chargeable, n: a.n + u.non_chargeable, l: a.l + u.leave }), { c: 0, n: 0, l: 0 });
const staff = D.staff;
const officeTot = staff.reduce((a, s) => a + s.office, 0);
const deptStaffTot = staff.reduce((a, s) => a + s.departmental, 0);

/* ============================================================== 1. title */
{
  const s = pres.addSlide({ masterName: 'CTS_DARK' });
  [{ x: 6.95, y: 0.30, d: 2.35, c: C.tealLt }, { x: 8.00, y: 1.30, d: 1.95, c: C.blue }, { x: 7.20, y: 2.25, d: 2.45, c: C.red }]
    .forEach((o, i) => s.addShape('ellipse', { x: o.x, y: o.y, w: o.d, h: o.d, fill: { color: o.c, transparency: 18 }, line: { type: 'none' }, objectName: `circle ${i}` }));
  s.addText('cts', { x: 0.55, y: 0.45, w: 2, h: 0.7, isTextBox: true, margin: 0, fontSize: 40, bold: true, color: C.white });
  s.addText('Seamless AV', { x: 0.62, y: 1.05, w: 2, h: 0.25, isTextBox: true, margin: 0, fontSize: 9.5, color: C.grey });
  s.addText('MONTHLY BUSINESS REVIEW', { x: 0.55, y: 2.00, w: 6, h: 0.3, isTextBox: true, margin: 0, fontSize: 10.5, bold: true, color: C.tealLt, charSpacing: 3 });
  s.addText(MONTH, { x: 0.55, y: 2.35, w: 6, h: 0.8, isTextBox: true, margin: 0, fontSize: 40, bold: true, color: C.white });
  s.addText('Results, departments, spending, customers, cash and utilisation', { x: 0.55, y: 3.30, w: 6, h: 0.4, isTextBox: true, margin: 0, fontSize: 13, color: C.lt2 || 'E3E6EA' });
  s.addText('Financial Controller', { x: 0.55, y: 4.55, w: 3, h: 0.25, isTextBox: true, margin: 0, fontSize: 9.5, color: C.grey });
  s.addNotes(`Monthly Business Review for ${MONTH}. Everything in this pack comes from the Financial Controller Pack, which ties to the Xero general ledger to the cent for every department and for the company.\n\nThe headline is that ${SHORT} beat budget on revenue by ${k(rev.month - rev.month_budget)} thousand and still lost money. Revenue was ${k(rev.month)} thousand against a budget of ${k(rev.month_budget)} thousand. EBITDA was ${k(eb.month)} thousand against a budgeted loss of ${k(eb.month_budget)} thousand.\n\nThe story of the month is margin and the overhead base, not sales.`);
}

/* ============================================================== 2. agenda */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Agenda', 'What we will cover');
  const items = [
    ['01', 'Financial performance', 'P&L summary, the three months to date, and every department against budget'],
    ['02', 'What we spend', 'Where the money goes: cost distribution, operating expense, labour and employment'],
    ['03', 'Customers and cash', 'Top ten customers, receivables and days sales outstanding'],
    ['04', 'People and utilisation', 'Chargeable hours by department and the profit they produce'],
    ['05', 'Budget and outlook', 'Revenue by month against budget, and what we do about it'],
  ];
  items.forEach(([n, t, sub], i) => {
    const y = 1.55 + i * 0.70;
    s.addShape('rect', { x: 0.55, y, w: 0.42, h: 0.42, fill: { color: C.navy }, line: { type: 'none' } });
    s.addText(n, { x: 0.55, y, w: 0.42, h: 0.42, isTextBox: true, margin: 0, fontSize: 11, bold: true, color: C.white, align: 'center', valign: 'middle' });
    s.addText(t, { x: 1.15, y: y - 0.02, w: 8.3, h: 0.26, isTextBox: true, margin: 0, fontSize: 13, bold: true, color: C.navy, valign: 'middle' });
    s.addText(sub, { x: 1.15, y: y + 0.22, w: 8.3, h: 0.24, isTextBox: true, margin: 0, fontSize: 10, color: C.muted, valign: 'middle' });
  });
  s.addNotes('Five sections. Financial performance first, because that is what the numbers say. What we spend second, because that is where the answer is this month. Then customers and cash, then people, then what we do about the budget.');
}

/* ------------------------------------------------- section divider helper */
function section(num, title, sub, notes) {
  const s = pres.addSlide({ masterName: 'CTS_DARK' });
  s.addShape('ellipse', { x: 7.05, y: 1.75, w: 2.9, h: 2.9, fill: { color: C.blue, transparency: 45 }, line: { type: 'none' }, objectName: 'divider circle' });
  s.addShape('rect', { x: 0.55, y: 2.05, w: 0.95, h: 0.95, fill: { color: C.tealLt }, line: { type: 'none' } });
  s.addText(num, { x: 0.55, y: 2.05, w: 0.95, h: 0.95, isTextBox: true, margin: 0, fontSize: 26, bold: true, color: C.navy, align: 'center', valign: 'middle' });
  s.addText('SECTION', { x: 0.55, y: 0.45, w: 4, h: 0.3, isTextBox: true, margin: 0, fontSize: 9.5, bold: true, color: C.tealLt, charSpacing: 2.2 });
  s.addText(title, { x: 1.75, y: 2.05, w: 6.5, h: 0.55, isTextBox: true, margin: 0, fontSize: 28, bold: true, color: C.white, valign: 'middle' });
  s.addText(sub, { x: 1.75, y: 2.62, w: 6.5, h: 0.38, isTextBox: true, margin: 0, fontSize: 12, color: C.grey, valign: 'middle' });
  if (notes) s.addNotes(notes);
  return s;
}

module.exports = { pres, D, P, rev, gp, oh, eb, np, depts, dTot, byRev, util, utilTot, staff, officeTot, deptStaffTot, ytdMargin, ohRun, breakEven, revRun, section, MONTH, SHORT };
