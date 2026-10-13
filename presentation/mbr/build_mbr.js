/*
 * Builds the CTS Monthly Business Review deck on the CTS slide masters.
 *   node presentation/mbr/build_mbr.js
 *
 * All figures come from mbr_data.json. Replace the numbers there each month
 * and rebuild; the layout, charts and commentary panels stay the same. The
 * data file ships with ILLUSTRATIVE figures.
 *
 * Every content slide has a commentary panel on the right (type over the
 * prompt) and a speaker-notes stub, so there is a place to comment on each
 * sheet.
 */
const path = require('path');
const cts = require('../cts_masters');

const { C, FONT, M } = cts;
const D = require('./mbr_data.json');
const pres = cts.newDeck(`CTS Monthly Business Review ${D.month}`);

// ---------------------------------------------------------------------------
// Geometry: content on the left, commentary panel on the right
// ---------------------------------------------------------------------------
const TOP = 1.6;          // first content row
const BOTTOM = 5.05;      // last content row
const CONTENT_W = 6.1;    // M .. 6.65
const PANEL_X = 6.85;
const PANEL_W = 9.45 - PANEL_X;

// ---------------------------------------------------------------------------
// Formatting
// ---------------------------------------------------------------------------
const money = (v) => (v == null ? '' : v < 0 ? `(${Math.abs(Math.round(v)).toLocaleString('en-AU')})` : Math.round(v).toLocaleString('en-AU'));
const pct = (v, dp = 1) => (v == null || !isFinite(v) ? '' : `${(v * 100).toFixed(dp)}%`);
const signed = (v) => (v == null ? '' : (v > 0 ? '+' : v < 0 ? '-' : '') + Math.abs(Math.round(v)).toLocaleString('en-AU'));
const signedPct = (v, dp = 1) => { if (v == null || !isFinite(v)) return ''; const r = Number((v * 100).toFixed(dp)); return `${r > 0 ? '+' : ''}${r.toFixed(dp)}%`; };
const sum = (arr, k) => arr.reduce((a, r) => a + (typeof k === 'function' ? k(r) : r[k]), 0);
const varColor = (v) => (v > 0 ? C.TEAL : v < 0 ? C.BURGUNDY : C.BLACK);

// ---------------------------------------------------------------------------
// Building blocks
// ---------------------------------------------------------------------------
function header(slide, section, title) {
  slide.addText(section, { placeholder: 'section' });
  slide.addText(title, { placeholder: 'title' });
}

// Commentary panel: the place to comment on each sheet. Type over the prompt.
function commentary(slide, prompts, { y = TOP, h = BOTTOM - TOP } = {}) {
  slide.addShape(pres.ShapeType.rect, { x: PANEL_X, y, w: PANEL_W, h, fill: { color: C.MINT_TINT }, line: { color: C.MINT_TINT, width: 0 } });
  slide.addText('COMMENTARY', { x: PANEL_X + 0.2, y: y + 0.18, w: PANEL_W - 0.4, h: 0.25, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 9, bold: true, color: C.TEAL, charSpacing: 2 });
  slide.addText(
    prompts.map((t, i) => ({ text: t, options: { bullet: { indent: 10 }, breakLine: i < prompts.length - 1 } })),
    { x: PANEL_X + 0.2, y: y + 0.48, w: PANEL_W - 0.4, h: h - 0.6, margin: 0, isTextBox: true, valign: 'top', fontFace: FONT, fontSize: 10.5, color: C.GREY, paraSpaceAfter: 6 },
  );
  slide.addNotes('Commentary: ' + prompts.join(' | '));
}

// A tinted tile with a big figure and a label.
function kpi(slide, { x, y, w, h = 0.95, value, label, color = C.BLACK, fill = C.LIGHT }) {
  slide.addShape(pres.ShapeType.rect, { x, y, w, h, fill: { color: fill }, line: { color: fill, width: 0 } });
  slide.addText(value, { x: x + 0.15, y: y + 0.1, w: w - 0.3, h: h * 0.52, margin: 0, isTextBox: true, valign: 'bottom', fontFace: FONT, fontSize: 22, bold: true, color });
  slide.addText(label, { x: x + 0.15, y: y + h * 0.62, w: w - 0.3, h: h * 0.34, margin: 0, isTextBox: true, valign: 'top', fontFace: FONT, fontSize: 9, color: C.GREY });
}

// Table with a charcoal header row, right-aligned figures, shaded total rows.
function table(slide, { x = M, y, w = CONTENT_W, colW, head, rows, fontSize = 9.5, rowH = 0.27 }) {
  const headRow = head.map((h, i) => ({
    text: h, options: { bold: true, color: C.WHITE, fill: { color: C.BLACK }, align: i === 0 ? 'left' : 'right', fontSize: fontSize - 0.5, valign: 'middle' },
  }));
  const body = rows.map((r) => r.cells.map((cell, i) => {
    const c = typeof cell === 'object' && cell !== null && 'text' in cell ? cell : { text: cell };
    return {
      text: c.text,
      options: {
        bold: !!r.total || !!c.bold, color: c.color || C.BLACK, align: i === 0 ? 'left' : 'right', valign: 'middle',
        fill: r.total ? { color: C.LIGHT } : r.fill ? { color: r.fill } : undefined,
        fontSize: r.small ? fontSize - 1 : fontSize,
      },
    };
  }));
  slide.addTable([headRow, ...body], {
    x, y, w, colW, rowH, fontFace: FONT, fontSize, color: C.BLACK, margin: [0.03, 0.06, 0.03, 0.06],
    border: [{ type: 'none' }, { type: 'none' }, { type: 'solid', pt: 0.5, color: C.LINE }, { type: 'none' }],
  });
  return y + rowH * (rows.length + 1);
}

const chartFrame = {
  catAxisLabelColor: C.GREY, catAxisLabelFontSize: 9, catAxisLabelFontFace: FONT, catGridLine: { style: 'none' }, catAxisLineShow: false,
  valAxisLabelColor: C.GREY, valAxisLabelFontSize: 9, valAxisLabelFontFace: FONT, valGridLine: { color: C.LINE, size: 0.5 }, valAxisLineShow: false,
  legendFontFace: FONT, legendFontSize: 9, legendColor: C.GREY,
  dataLabelFontFace: FONT, dataLabelFontSize: 8, dataLabelColor: C.GREY,
};

function section(master, number, title, subtitle) {
  const s = pres.addSlide({ masterName: master });
  s.addText(number, { placeholder: 'number' });
  s.addText(title, { placeholder: 'title' });
  s.addText(subtitle, { placeholder: 'subtitle' });
}

// ---------------------------------------------------------------------------
// Derived figures
// ---------------------------------------------------------------------------
const depts = D.departments;
const opDepts = depts.filter((d) => d.revenue > 0);
const totalRevenue = sum(depts, 'revenue');
const totalYtd = sum(D.topCustomers, 'ytd');
const ytdRevenue = D.plSummary.rows.find((r) => r.line === 'Revenue').ytd;
const util = D.utilisation.rows.map((r) => {
  const capacity = r.fte * D.workingDays * D.hoursPerDay;
  return { ...r, capacity, pct: r.chargeable / capacity };
});
const utilTotal = {
  dept: 'CTS', fte: sum(util, 'fte'), chargeable: sum(util, 'chargeable'), nonChargeable: sum(util, 'nonChargeable'), leave: sum(util, 'leave'), capacity: sum(util, 'capacity'),
};
utilTotal.pct = utilTotal.chargeable / utilTotal.capacity;
const labour = D.labourProfit.rows.map((r) => {
  const u = util.find((x) => x.dept === r.dept);
  const revenue = (u.chargeable * r.chargeRate) / 1000;
  const cost = (u.capacity * r.costRate) / 1000;
  const profit = revenue - cost;
  const uplift = (0.05 * u.capacity * r.chargeRate) / 1000;
  return { ...r, hours: u.chargeable, revenue, cost, profit, perHour: (profit * 1000) / u.chargeable, uplift, pct: u.pct };
});

// ===========================================================================
// Slides
// ===========================================================================

// 1. Cover
{
  const s = pres.addSlide({ masterName: 'CTS Cover (dark)' });
  s.addText('MONTHLY BUSINESS REVIEW', { placeholder: 'eyebrow' });
  s.addText(D.month, { placeholder: 'title' });
  s.addText('Results, risk revenue, customers, cash and utilisation', { placeholder: 'subtitle' });
  s.addText(D.presenter, { placeholder: 'presenter' });
  s.addText(D.month, { placeholder: 'date' });
  s.addNotes('Monthly Business Review. All figures in the data file are illustrative until replaced with the month\'s actuals.');
}

// 2. Agenda
{
  const s = pres.addSlide({ masterName: 'CTS Header and Text' });
  header(s, 'AGENDA', 'What we will cover');
  const items = [
    'Financial performance: P&L summary, P&L per department, risk revenue profile',
    'Customers and cash: top 10 customers, DSO and receivables, top customers year on year',
    'People and utilisation: utilisation by department, profit based on utilisation',
    'Summary and actions',
  ];
  s.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1 } })), { placeholder: 'body' });
}

// 3. Section: financial performance
section('CTS Section (navy)', '01', 'Financial performance', `P&L for ${D.month} and ${D.ytdLabel}`);

// 4. P&L summary
{
  const s = pres.addSlide({ masterName: 'CTS Header Only' });
  header(s, 'FINANCIAL PERFORMANCE', 'P&L summary');
  const rows = D.plSummary.rows.map((r) => {
    const f = r.pct ? (v) => pct(v) : money;
    const mv = r.pct ? null : r.month - r.monthBudget;
    const yv = r.pct ? null : r.ytd - r.ytdBudget;
    return {
      total: r.total, small: r.pct,
      cells: [r.line, f(r.month), f(r.monthBudget), r.pct ? signedPct(r.month - r.monthBudget) : { text: signed(mv), color: varColor(mv) },
        f(r.ytd), f(r.ytdBudget), r.pct ? signedPct(r.ytd - r.ytdBudget) : { text: signed(yv), color: varColor(yv) }, f(r.priorYearMonth)],
    };
  });
  const end = table(s, {
    y: TOP, colW: [1.5, 0.66, 0.66, 0.66, 0.66, 0.66, 0.66, 0.64],
    head: ['A$ thousands', D.monthShort, 'Budget', 'Var', 'YTD', 'Budget', 'Var', `PY ${D.monthShort}`], rows,
  });
  const t = D.plSummary.trend;
  s.addText('Revenue and EBITDA, last six months', { x: M, y: end + 0.12, w: CONTENT_W, h: 0.22, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 9, bold: true, color: C.GREY });
  s.addChart(pres.ChartType.bar, [
    { name: 'Revenue', labels: t.labels, values: t.revenue },
    { name: 'EBITDA', labels: t.labels, values: t.ebitda },
  ], { x: M, y: end + 0.32, w: CONTENT_W, h: BOTTOM - end - 0.32, barDir: 'col', barGrouping: 'clustered', chartColors: [C.NAVY, C.MINT], showLegend: true, legendPos: 'r', showValue: false, ...chartFrame });
  commentary(s, ['What drove the month against budget', 'YTD position and the run-rate to full year', 'One-offs or timing to call out']);
}

// 5. P&L per department (table)
{
  const s = pres.addSlide({ masterName: 'CTS Header Only' });
  header(s, 'FINANCIAL PERFORMANCE', `P&L per department, ${D.monthShort}`);
  const rows = depts.map((d) => {
    const v = d.revenue - d.revenueBudget;
    const contrib = d.grossProfit - d.overheads;
    return { cells: [d.name, money(d.revenue), money(d.revenueBudget), { text: signed(v), color: varColor(v) }, money(d.grossProfit), d.revenue ? pct(d.grossProfit / d.revenue) : '', money(-d.overheads), { text: money(contrib), color: contrib < 0 ? C.BURGUNDY : C.BLACK }] };
  });
  const tot = {
    revenue: totalRevenue, budget: sum(depts, 'revenueBudget'), gp: sum(depts, 'grossProfit'), oh: sum(depts, 'overheads'),
  };
  rows.push({ total: true, cells: ['CTS', money(tot.revenue), money(tot.budget), { text: signed(tot.revenue - tot.budget), color: varColor(tot.revenue - tot.budget) }, money(tot.gp), pct(tot.gp / tot.revenue), money(-tot.oh), money(tot.gp - tot.oh)] });
  const end = table(s, {
    y: TOP, colW: [1.2, 0.7, 0.7, 0.6, 0.78, 0.6, 0.7, 0.82],
    head: ['A$ thousands', 'Revenue', 'Budget', 'Var', 'Gross profit', 'GM %', 'Overheads', 'Contribution'], rows,
  });
  s.addText('Overheads are the department\'s direct overheads; Admin carries the shared overheads. Contribution is gross profit less overheads, and the CTS row is EBITDA.',
    { x: M, y: end + 0.12, w: CONTENT_W, h: 0.5, margin: 0, isTextBox: true, valign: 'top', fontFace: FONT, fontSize: 9, color: C.GREY });
  commentary(s, ['Which departments beat or missed budget, and why', 'Margin movement by department', 'Overhead changes worth noting']);
}

// 6. P&L per department (chart)
{
  const s = pres.addSlide({ masterName: 'CTS Header Only' });
  header(s, 'FINANCIAL PERFORMANCE', 'Revenue and contribution by department');
  const labels = opDepts.map((d) => d.name);
  s.addChart(pres.ChartType.bar, [
    { name: 'Revenue', labels, values: opDepts.map((d) => d.revenue) },
    { name: 'Budget revenue', labels, values: opDepts.map((d) => d.revenueBudget) },
    { name: 'Contribution', labels, values: opDepts.map((d) => d.grossProfit - d.overheads) },
  ], { x: M, y: TOP, w: CONTENT_W, h: BOTTOM - TOP, barDir: 'col', barGrouping: 'clustered', chartColors: [C.NAVY, C.SOFT, C.MINT], showLegend: true, legendPos: 'b', showValue: true, dataLabelPosition: 'outEnd', ...chartFrame });
  commentary(s, ['Where the mix of revenue is shifting', 'Departments running below contribution expectations', 'What changes next month']);
}

// 7. Risk revenue profile
{
  const s = pres.addSlide({ masterName: 'CTS Header Only' });
  header(s, 'FINANCIAL PERFORMANCE', 'Risk revenue profile');
  const contract = sum(depts, 'contract');
  const nonContract = sum(depts, 'nonContract');
  kpi(s, { x: M, y: TOP, w: 1.9, h: 0.7, value: pct(contract / totalRevenue, 0), label: 'Contracted revenue, secured', color: C.TEAL });
  kpi(s, { x: M + 2.05, y: TOP, w: 1.9, h: 0.7, value: pct(nonContract / totalRevenue, 0), label: 'Non-contract revenue, at risk', color: C.BURGUNDY });
  kpi(s, { x: M + 4.1, y: TOP, w: 2.0, h: 0.7, value: String(D.riskFlags.rows.filter((r) => r.risk !== 'Low').length), label: 'Revenue lines flagged Medium or High', color: C.BLACK });
  const labels = opDepts.map((d) => d.name);
  s.addChart(pres.ChartType.bar, [
    { name: 'Contract', labels, values: opDepts.map((d) => d.contract) },
    { name: 'Non-contract', labels, values: opDepts.map((d) => d.nonContract) },
  ], { x: M, y: TOP + 0.8, w: CONTENT_W, h: 1.3, barDir: 'bar', barGrouping: 'stacked', chartColors: [C.TEAL, C.ROSE], showLegend: true, legendPos: 'r', showValue: true, dataLabelPosition: 'ctr', ...chartFrame, dataLabelColor: C.WHITE, catAxisLabelFontSize: 8, valAxisLabelFontSize: 8 });
  const riskColor = { High: C.BURGUNDY, Medium: C.NAVY, Low: C.GREY };
  table(s, {
    y: TOP + 2.2, rowH: 0.205, fontSize: 8.5, colW: [2.5, 0.72, 0.72, 0.62, 0.72, 0.82],
    head: ['Revenue line vs budget, A$ thousands', 'Actual', 'Budget', 'Var', 'Var %', 'Risk'],
    rows: D.riskFlags.rows.map((r) => ({ cells: [r.line, money(r.actual), money(r.budget), { text: signed(r.var), color: varColor(r.var) }, signedPct(r.varPct), { text: r.risk, bold: true, color: riskColor[r.risk] }] })),
  });
  commentary(s, ['How much of next quarter is already contracted', 'Non-contract revenue: pipeline cover and conversion', 'Flagged lines: cause and action']);
}

// 8. Section: customers and cash
section('CTS Section (navy)', '02', 'Customers and cash', 'Top customers, receivables and days sales outstanding');

// 9. Top 10 customers
{
  const s = pres.addSlide({ masterName: 'CTS Header Only' });
  header(s, 'CUSTOMERS AND CASH', 'Top 10 customers');
  const rows = D.topCustomers.map((c) => {
    const yoy = c.lastYtd ? (c.ytd - c.lastYtd) / c.lastYtd : null;
    return { cells: [`${c.rank}.  ${c.client}`, money(c.month), money(c.ytd), pct(c.ytd / ytdRevenue, 0), money(c.lastYtd), c.lastYtd ? { text: signedPct(yoy, 0), color: varColor(yoy) } : { text: 'New', color: C.TEAL }] };
  });
  const lastTotal = sum(D.topCustomers, 'lastYtd');
  rows.push({ total: true, cells: ['Top 10', money(sum(D.topCustomers, 'month')), money(totalYtd), pct(totalYtd / ytdRevenue, 0), money(lastTotal), { text: signedPct((totalYtd - lastTotal) / lastTotal, 0), color: varColor(totalYtd - lastTotal) }] });
  const end = table(s, {
    y: TOP, rowH: 0.255, colW: [1.9, 0.8, 0.8, 0.9, 0.9, 0.8],
    head: ['Client, A$ thousands', D.monthShort, 'YTD', 'Share of YTD', `${D.priorFy} YTD`, 'YoY'], rows,
  });
  s.addText(`Top 10 customers are ${pct(totalYtd / ytdRevenue, 0)} of ${D.ytdLabel} revenue. Client names come from the Clients sheet in the Financial Controller Pack.`,
    { x: M, y: end + 0.1, w: CONTENT_W, h: 0.4, margin: 0, isTextBox: true, valign: 'top', fontFace: FONT, fontSize: 9, color: C.GREY });
  commentary(s, ['Movements in the top 10 and why', 'Concentration risk: how dependent we are on the top three', 'New or lost customers']);
}

// 10. DSO and receivables
{
  const s = pres.addSlide({ masterName: 'CTS Header Only' });
  header(s, 'CUSTOMERS AND CASH', 'Receivables and days sales outstanding');
  const dso = D.dso;
  const overdue60 = dso.aged.filter((a) => a.bucket.startsWith('61') || a.bucket.startsWith('Over')).reduce((a, b) => a + b.amount, 0);
  kpi(s, { x: M, y: TOP, w: 1.9, h: 0.85, value: `${dso.days} days`, label: `DSO, target ${dso.target} days`, color: dso.days > dso.target ? C.BURGUNDY : C.TEAL });
  kpi(s, { x: M + 2.05, y: TOP, w: 1.9, h: 0.85, value: `$${(dso.receivables / 1000).toFixed(2)}m`, label: 'Trade receivables at month end', color: C.BLACK });
  kpi(s, { x: M + 4.1, y: TOP, w: 2.0, h: 0.85, value: `$${overdue60}k`, label: 'Over 60 days', color: overdue60 > 0 ? C.BURGUNDY : C.TEAL });
  s.addText('DSO, last six months', { x: M, y: TOP + 1.0, w: 3.2, h: 0.22, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 9, bold: true, color: C.GREY });
  s.addChart(pres.ChartType.line, [
    { name: 'DSO', labels: dso.trend.labels, values: dso.trend.days },
    { name: 'Target', labels: dso.trend.labels, values: dso.trend.labels.map(() => dso.target) },
  ], { x: M, y: TOP + 1.2, w: 3.2, h: BOTTOM - TOP - 1.2, chartColors: [C.NAVY, C.ROSE], lineSize: 2, lineDataSymbol: 'circle', lineDataSymbolSize: 5, showLegend: true, legendPos: 'b', showValue: false, ...chartFrame });
  s.addText('Aged receivables, A$ thousands', { x: M + 3.4, y: TOP + 1.0, w: 2.7, h: 0.22, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 9, bold: true, color: C.GREY });
  table(s, {
    x: M + 3.4, y: TOP + 1.25, w: 2.7, rowH: 0.26, fontSize: 9, colW: [1.5, 0.6, 0.6],
    head: ['Bucket', 'A$k', '%'],
    rows: [...dso.aged.map((a) => ({ cells: [a.bucket, money(a.amount), pct(a.amount / dso.receivables, 0)] })), { total: true, cells: ['Total', money(dso.receivables), '100%'] }],
  });
  s.addText('DSO = receivables / (last three months revenue / 92 days).', { x: M + 3.4, y: BOTTOM - 0.3, w: 2.7, h: 0.3, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 8, color: C.GREY });
  commentary(s, ['Largest overdue accounts and collection status', 'Why DSO moved this month', 'Cash actions before next month end']);
}

// 11. Top customers year on year
{
  const s = pres.addSlide({ masterName: 'CTS Header Only' });
  header(s, 'CUSTOMERS AND CASH', `Top customers, ${D.fy} vs ${D.priorFy} year to date`);
  const labels = D.topCustomers.map((c) => c.client);
  s.addChart(pres.ChartType.bar, [
    { name: `${D.fy} YTD`, labels, values: D.topCustomers.map((c) => c.ytd) },
    { name: `${D.priorFy} YTD`, labels, values: D.topCustomers.map((c) => c.lastYtd) },
  ], { x: M, y: TOP, w: CONTENT_W, h: BOTTOM - TOP - 0.35, barDir: 'col', barGrouping: 'clustered', chartColors: [C.NAVY, C.SOFT], showLegend: true, legendPos: 'b', showValue: false, ...chartFrame });
  s.addText('A$ thousands, same period last year. Customers ranked by current year to date.', { x: M, y: BOTTOM - 0.3, w: CONTENT_W, h: 0.3, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 8, color: C.GREY });
  commentary(s, ['Customers growing year on year and what is driving it', 'Customers in decline and the recovery plan', 'New logos replacing lost revenue']);
}

// 12. Section: people and utilisation
section('CTS Section (navy)', '03', 'People and utilisation', 'Chargeable hours, utilisation and the profit they produce');

// 13. Utilisation
{
  const s = pres.addSlide({ masterName: 'CTS Header Only' });
  header(s, 'PEOPLE AND UTILISATION', `Utilisation by department, ${D.monthShort}`);
  const target = D.utilisation.target;
  const rows = util.map((u) => ({ cells: [u.dept, String(u.fte), money(u.chargeable), money(u.nonChargeable), money(u.leave), { text: pct(u.pct), bold: true, color: u.pct >= target ? C.TEAL : C.BURGUNDY }, pct(u.priorMonth), { text: signedPct(u.pct - target), color: varColor(u.pct - target) }] }));
  rows.push({ total: true, cells: ['CTS', String(utilTotal.fte), money(utilTotal.chargeable), money(utilTotal.nonChargeable), money(utilTotal.leave), { text: pct(utilTotal.pct), color: utilTotal.pct >= target ? C.TEAL : C.BURGUNDY }, '', { text: signedPct(utilTotal.pct - target), color: varColor(utilTotal.pct - target) }] });
  const end = table(s, {
    y: TOP, rowH: 0.26, colW: [1.15, 0.5, 0.85, 0.95, 0.65, 0.7, 0.65, 0.65],
    head: ['Hours', 'FTE', 'Chargeable', 'Non-chargeable', 'Leave', 'Util %', 'Prior', 'vs target'], rows,
  });
  s.addChart(pres.ChartType.bar, [
    { name: 'Utilisation', labels: util.map((u) => u.dept), values: util.map((u) => Math.round(u.pct * 1000) / 10) },
    { name: 'Target', labels: util.map((u) => u.dept), values: util.map(() => target * 100) },
  ], { x: M, y: end + 0.15, w: CONTENT_W, h: BOTTOM - end - 0.15, barDir: 'col', barGrouping: 'clustered', chartColors: [C.MINT, C.SOFT], showLegend: true, legendPos: 'r', showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.0', valAxisMaxVal: 100, ...chartFrame });
  s.addNotes(`Utilisation % = chargeable hours / (FTE x ${D.workingDays} working days x ${D.hoursPerDay} hours). Target ${pct(target, 0)}.`);
  commentary(s, ['Departments under target and the reason: bench time, leave, training', 'Overtime or casual hours masking low utilisation', 'Hiring or roster changes planned']);
}

// 14. Profit based on utilisation
{
  const s = pres.addSlide({ masterName: 'CTS Header Only' });
  header(s, 'PEOPLE AND UTILISATION', 'Profit based on utilisation');
  const rows = labour.map((l) => ({ cells: [l.dept, pct(l.pct), money(l.hours), `$${l.chargeRate}`, money(l.revenue), money(-l.cost), { text: money(l.profit), bold: true, color: l.profit < 0 ? C.BURGUNDY : C.BLACK }, `$${Math.round(l.perHour)}`, { text: `+${money(l.uplift)}`, color: C.TEAL }] }));
  const tot = { hours: sum(labour, 'hours'), revenue: sum(labour, 'revenue'), cost: sum(labour, 'cost'), profit: sum(labour, 'profit'), uplift: sum(labour, 'uplift') };
  rows.push({ total: true, cells: ['CTS', pct(utilTotal.pct), money(tot.hours), '', money(tot.revenue), money(-tot.cost), money(tot.profit), `$${Math.round((tot.profit * 1000) / tot.hours)}`, { text: `+${money(tot.uplift)}`, color: C.TEAL }] });
  const end = table(s, {
    y: TOP, rowH: 0.26, fontSize: 9, colW: [1.0, 0.6, 0.7, 0.55, 0.8, 0.75, 0.7, 0.5, 0.5],
    head: ['A$ thousands', 'Util %', 'Charge hrs', 'Rate', 'Labour rev', 'Labour cost', 'Labour profit', 'Per hr', '+5 pts'], rows,
  });
  kpi(s, { x: M, y: end + 0.15, w: 2.9, h: BOTTOM - end - 0.15, value: `+$${money(tot.uplift)}k a month`, label: 'Five more points of utilisation at current charge rates, before any extra cost', color: C.TEAL, fill: C.MINT_TINT });
  s.addText('Labour revenue = chargeable hours x average charge rate. Labour cost = FTE x working hours x loaded cost per hour, so it includes the cost of non-chargeable and leave hours. Per hr = labour profit per chargeable hour. Rates come from the pack\'s Setup and Utilisation sheets.',
    { x: M + 3.05, y: end + 0.15, w: CONTENT_W - 3.05, h: BOTTOM - end - 0.15, margin: 0, isTextBox: true, valign: 'top', fontFace: FONT, fontSize: 8.5, color: C.GREY });
  commentary(s, ['Where a point of utilisation is worth the most', 'Rate pressure: charge rate vs loaded cost', 'Actions to lift chargeable hours']);
}

// 15. Section: summary and actions
section('CTS Section (light)', '04', 'Summary and actions', 'The commentary from each sheet, with an owner and a date');

// 16. Actions
{
  const s = pres.addSlide({ masterName: 'CTS Header Only' });
  header(s, 'SUMMARY AND ACTIONS', 'Commentary and actions from this review');
  const rows = D.actions.map((a) => ({ cells: [a.area, a.comment || 'Type the commentary agreed in the meeting', a.owner || 'Owner', a.due || 'Date'] }));
  const headRow = ['Area', 'Commentary and agreed action', 'Owner', 'Due'].map((h, i) => ({ text: h, options: { bold: true, color: C.WHITE, fill: { color: C.BLACK }, align: 'left', fontSize: 9.5, valign: 'middle' } }));
  s.addTable([headRow, ...rows.map((r, i) => r.cells.map((t, j) => ({ text: t, options: { align: 'left', valign: 'top', color: j === 0 ? C.BLACK : C.GREY, bold: j === 0 } })))], {
    x: M, y: TOP, w: 9.45 - M, colW: [1.5, 5.4, 1.0, 1.0], rowH: [0.3, 0.75, 0.75, 0.75, 0.75], fontFace: FONT, fontSize: 10, margin: [0.06, 0.08, 0.06, 0.08],
    border: [{ type: 'none' }, { type: 'none' }, { type: 'solid', pt: 0.5, color: C.LINE }, { type: 'none' }],
  });
  s.addNotes('Carry the commentary from each sheet into this table during the review so the actions and owners are recorded in one place.');
}

cts.writeDeck(pres, path.join(__dirname, `CTS Monthly Business Review - ${D.month}.pptx`))
  .catch((err) => { console.error(err); process.exit(1); });
