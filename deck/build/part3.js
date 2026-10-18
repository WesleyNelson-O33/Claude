const B = require('./build.js');
const L = require('./lib.js');
const { C, DEPT } = require('./theme.js');
const { head, commentary, note, tile, row, headRow, CHART, money, k, pct, signed, varColour } = L;
const { pres, D, P, rev, gp, oh, eb, np, depts, dTot, byRev, util, utilTot, staff, ytdMargin, ohRun, breakEven, revRun, section, MONTH, SHORT } = B;

/* ===================================================== 14. section 03 */
section('03', 'Customers and cash', 'Who pays us, how much, and how slowly',
  `Revenue is concentrated and cash is slow. Days sales outstanding is ${D.dso.days.toFixed(1)} against a ${D.dso.target} day target, which is ${money(D.dso.locked)} of cash sitting in the debtor book above where it should be.`);

/* ===================================================== 15. top ten customers */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Customers and cash', 'Top ten customers, year to date');
  const t = D.top10, tt = D.top10_total;
  const body = t.map((x) => row([
    { text: `${x.rank}. ${String(x.name).slice(0, 26)}`, align: 'left' },
    money(x.month), money(x.ytd), pct(x.share),
  ]));
  body.push(row([{ text: 'Top ten together', align: 'left' }, money(tt.top_ten_month), money(tt.top_ten_ytd), pct(tt.top_ten_share)], { bold: true, fill: C.tile }));
  body.push(row([{ text: 'Everyone else', align: 'left' }, money(tt.rest_month), money(tt.rest_ytd), pct(tt.rest_share)]));
  s.addTable([headRow(['Customer, A$', SHORT, 'Year to date', 'Share of YTD']), ...body], {
    x: 0.55, y: 1.58, w: 6.10, colW: [2.65, 1.12, 1.25, 1.08],
    border: { type: 'solid', color: C.rule, pt: 0.5 }, autoPage: false,
  });
  note(s, 'Earned revenue off the Xero general ledger, excluding GST, 1 July to 30 September 2026. Ties to total income on the P&L.',
    { x: 0.55, y: 4.90, w: 6.10, h: 0.3 });
  const top3 = t.slice(0, 3).reduce((a, x) => a + x.ytd, 0);
  commentary(s, [
    `The top ten are ${pct(tt.top_ten_share)} of revenue. The top three alone are ${pct(top3 / rev.ytd)}.`,
    `${t[0].name} at ${money(t[0].ytd)} is ${pct(t[0].share)} of the year. Losing it costs more than the whole EBITDA plan.`,
    `${t[1].name} came from nothing to ${money(t[1].ytd)} in a quarter, almost all of it in September.`,
    `Below the top five the book drops away fast. Number ten is ${money(t[9].ytd)}.`,
    `Recommend a named retention plan on the top three before the half year.`,
  ]);
  s.addNotes(`Top ten customers, year to date.\n\n` + t.map((x) => `${x.rank}. ${x.name}: ${money(x.ytd)} year to date at ${pct(x.share)} of revenue, ${money(x.month)} in ${MONTH}.`).join('\n') + `\n\nThe top ten are ${money(tt.top_ten_ytd)} or ${pct(tt.top_ten_share)} of revenue. Everyone else is ${money(tt.rest_ytd)}.\n\nThe concentration risk is real and it is getting worse, not better. ${t[0].name} and ${t[1].name} together are ${pct((t[0].ytd + t[1].ytd) / rev.ytd)} of the year to date. ${t[1].name} is a single project, so that share will not repeat next quarter.\n\nThe recommendation is a retention plan with a named owner for the top three, and a business development target set on the second ten rather than the first.`);
}

/* ===================================================== 16. top ten chart */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Customers and cash', 'Top ten customers by revenue');
  const t = D.top10;
  s.addChart('bar', [{ name: 'Year to date', labels: t.map((x) => String(x.name).slice(0, 22)), values: t.map((x) => Math.round(x.ytd / 1000)) }], {
    x: 0.50, y: 1.60, w: 6.20, h: 3.30, barDir: 'bar', barGapWidthPct: 38,
    ...CHART, chartColors: [C.teal],
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 8.5,
    catAxisOrientation: 'maxMin', valAxisHidden: true, valGridLine: { style: 'none' },
  });
  note(s, 'A$ thousands, 1 July to 30 September 2026.', { x: 0.55, y: 4.92, w: 6.10, h: 0.3 });
  commentary(s, [
    `The shape of this chart is the risk. Two bars carry the quarter.`,
    `${t[0].name} is contract and recurring. That is the good concentration.`,
    `${t[1].name} is project. That is the concentration that disappears.`,
    `Strip the top two and the remaining eight are ${money(t.slice(2).reduce((a, x) => a + x.ytd, 0))}, under two months of overhead.`,
    `The business development target belongs on the middle of this chart, not the top.`,
  ]);
  s.addNotes(`The same ten customers as a chart, so the distribution is visible.\n\nTwo customers dominate. ${t[0].name} at ${money(t[0].ytd)} and ${t[1].name} at ${money(t[1].ytd)}. The remaining eight together are ${money(t.slice(2).reduce((a, x) => a + x.ytd, 0))}.\n\nThe distinction that matters is contract versus project. ${t[0].name} is contracted support income that renews. ${t[1].name} is a fit-out that completes. Treating them as the same kind of revenue is how a forecast goes wrong.`);
}

/* ===================================================== 17. receivables and DSO */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Customers and cash', 'Receivables and days sales outstanding');
  const d = D.dso;
  tile(s, { x: 0.55, y: 1.58, w: 1.82, h: 0.95, value: `${d.days.toFixed(0)} days`, label: `DSO, target ${d.target} days`, colour: C.bad });
  tile(s, { x: 2.52, y: 1.58, w: 1.82, h: 0.95, value: `$${(d.receivables / 1e6).toFixed(2)}m`, label: 'Trade receivables at month end', colour: C.navy });
  tile(s, { x: 4.49, y: 1.58, w: 1.82, h: 0.95, value: `$${Math.round(d.locked / 1000)}k`, label: 'Cash locked up above target', colour: C.bad });
  const ag = D.ageing.filter((x) => x.amount > 0);
  s.addChart('bar', [{ name: 'Ageing', labels: ag.map((x) => x.bucket), values: ag.map((x) => Math.round(x.amount / 1000)) }], {
    x: 0.50, y: 2.70, w: 2.95, h: 2.05, barDir: 'col', barGapWidthPct: 45,
    ...CHART, chartColors: [C.teal, C.tealLt, C.red, C.bad],
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 8,
    valAxisHidden: true, valGridLine: { style: 'none' },
  });
  s.addText('Aged receivables, A$ thousands', { x: 0.55, y: 4.72, w: 2.9, h: 0.22, isTextBox: true, margin: 0, fontSize: 8.5, color: C.muted });
  const deb = D.debtors.slice(0, 6);
  s.addTable([headRow(['Who owes the most', 'A$', 'Share']),
    ...deb.map((x) => row([{ text: String(x.name).slice(0, 24), align: 'left' }, money(x.amount), pct(x.share)])),
    row([{ text: 'All others', align: 'left' }, money(d.receivables - D.debtors.reduce((a, x) => a + x.amount, 0) + D.debtors.slice(6).reduce((a, x) => a + x.amount, 0)), ''], { bold: true, fill: C.tile }),
  ], { x: 3.60, y: 2.70, w: 3.05, colW: [1.55, 0.85, 0.65], border: { type: 'solid', color: C.rule, pt: 0.5 }, autoPage: false });
  note(s, 'DSO = receivables divided by average daily revenue over the last three months (92 days). Source: Xero aged receivables at 30 September 2026.',
    { x: 3.60, y: 4.72, w: 3.05, h: 0.4 });
  commentary(s, [
    `DSO is ${d.days.toFixed(1)} days against a ${d.target} day target. That is ${d.above.toFixed(1)} days of cash we are not holding.`,
    `Every day of DSO is ${money(d.receivables / d.days)} of cash. Getting to target releases ${money(d.locked)}.`,
    `${pct(d.overdue_share)} of the book is overdue, ${money(d.overdue)}.`,
    `${D.debtors[0].name} alone is ${pct(D.debtors[0].share)} of receivables at ${money(D.debtors[0].amount)}.`,
    `Chase the top two. Nothing else on this page moves the number.`,
  ]);
  s.addNotes(`Receivables and days sales outstanding at 30 September 2026.\n\nTrade receivables ${money(d.receivables)}. DSO ${d.days.toFixed(1)} days against a target of ${d.target}. At target the book would be ${money(d.at_target)}, so ${money(d.locked)} of cash is tied up above where it should be.\n\nAgeing: ` + ag.map((x) => `${x.bucket} ${money(x.amount)} at ${pct(x.share)}`).join(', ') + `. Overdue in total ${money(d.overdue)} or ${pct(d.overdue_share)}.\n\nThe largest debtors are ` + deb.slice(0, 3).map((x) => `${x.name} at ${money(x.amount)}`).join(', ') + `.\n\nThe practical action is narrow. Two customers hold the majority of the overdue balance. A targeted chase on those two is worth more than a general collections push, and it is the fastest cash available to the business this quarter.`);
}

/* ===================================================== 18. section 04 */
section('04', 'People and utilisation', 'Chargeable hours and the profit they produce',
  `The company billed ${utilTot.c.toLocaleString('en-AU', { maximumFractionDigits: 0 })} chargeable hours in ${MONTH} out of ${(utilTot.c + utilTot.n).toLocaleString('en-AU', { maximumFractionDigits: 0 })} worked. Where chargeable hours fall short of target, margin follows.`);

/* ===================================================== 19. utilisation */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'People and utilisation', `Utilisation by department, ${SHORT}`);
  const body = util.map((u) => row([
    { text: u.name, align: 'left' }, u.fte.toFixed(1),
    money(u.chargeable), money(u.non_chargeable), money(u.leave),
    { text: pct(u.util_pct), bold: true, color: u.util_pct >= 0.8 ? C.good : C.bad },
  ]));
  const tc = utilTot.c + utilTot.n;
  body.push(row([{ text: 'Five departments', align: 'left' }, util.reduce((a, u) => a + u.fte, 0).toFixed(1),
    money(utilTot.c), money(utilTot.n), money(utilTot.l),
    { text: pct(utilTot.c / tc), bold: true }], { bold: true, fill: C.tile }));
  s.addTable([headRow([MONTH, 'FTE', 'Chargeable', 'Non-chargeable', 'Leave', 'Chg %']), ...body], {
    x: 0.55, y: 1.60, w: 6.10, colW: [1.3, 0.72, 1.15, 1.32, 0.78, 0.83],
    border: { type: 'solid', color: C.rule, pt: 0.5 }, autoPage: false,
  });
  note(s, 'Hours are counted against the cost centre they were booked to. Chargeable per cent is chargeable hours over hours worked, with leave excluded. Admin hours are charged out through the overhead keys and are not in this table.',
    { x: 0.55, y: 3.55, w: 6.10, h: 0.45 });
  const sup = util[0], ev = util[1], con = util[4];
  commentary(s, [
    `Support shows ${money(sup.non_chargeable)} non-chargeable hours on ${money(sup.chargeable)} chargeable. That is not credible.`,
    `Until that coding is fixed the ${pct(sup.util_pct)} means nothing, and nor does the target.`,
    `Events is at ${pct(ev.util_pct)} with ${money(ev.leave)} hours of leave. Real capacity is sitting idle.`,
    `Consulting at ${pct(con.util_pct)} confirms there is no fee base to utilise.`,
    `Hours and revenue sit on the same basis, so the pack is consistent even where coding is wrong.`,
  ]);
  s.addNotes(`Utilisation by department for ${MONTH}.\n\n` + util.map((u) => `${u.name}: ${money(u.chargeable)} chargeable, ${money(u.non_chargeable)} non-chargeable, ${money(u.leave)} leave, ${pct(u.util_pct)} chargeable.`).join('\n') + `\n\nCompany total ${money(utilTot.c)} chargeable out of ${money(tc)} worked, ${pct(utilTot.c / tc)}.\n\nI need to be straight about the Support number. ${pct(sup.util_pct)} chargeable on only ${money(sup.non_chargeable)} non-chargeable hours in a whole month is not believable for a team of that size. Non-chargeable time is being booked to chargeable cost centres. The hours total is right, the split is not, and that makes the Support utilisation figure unusable until timesheet coding is corrected.\n\nEvents at ${pct(ev.util_pct)} is the real operational issue. ${money(ev.leave)} hours of leave in the month on a small team takes out most of the capacity.`);
}

/* ===================================================== 20. profit from utilisation */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'People and utilisation', 'Profit based on utilisation');
  const rows = util.map((u, i) => {
    const dep = depts.find((x) => x.name === u.name);
    const revPerHr = u.chargeable ? dep.revenue / u.chargeable : 0;
    const labPerHr = u.chargeable ? Math.abs(dep.cos) / u.chargeable : 0;
    return { name: u.name, chg: u.chargeable, pctv: u.util_pct, revPerHr, labPerHr, net: revPerHr - labPerHr, five: revPerHr * u.chargeable * 0.05 };
  });
  const meaningful = (r) => r.chg >= 50;      // below that the ratio is noise, not a rate
  const body = rows.map((r) => row(meaningful(r) ? [
    { text: r.name, align: 'left' }, pct(r.pctv), money(r.chg),
    '$' + money(r.revPerHr), '$' + money(r.labPerHr),
    { text: '$' + money(r.net), bold: true, color: r.net >= 0 ? C.good : C.bad },
    money(r.five),
  ] : [
    { text: r.name, align: 'left' }, pct(r.pctv), money(r.chg),
    { text: 'n/a', color: C.muted }, { text: 'n/a', color: C.muted },
    { text: 'n/a', color: C.muted }, { text: 'n/a', color: C.muted },
  ]));
  s.addTable([headRow([MONTH, 'Chg %', 'Chg hrs', 'Revenue / hr', 'Cost / hr', 'Margin / hr', '+5 pts worth']), ...body], {
    x: 0.55, y: 1.60, w: 6.10, colW: [1.18, 0.66, 0.78, 1.05, 0.85, 0.95, 0.63],
    border: { type: 'solid', color: C.rule, pt: 0.5 }, autoPage: false,
  });
  const gain = rows[0].five + rows[1].five;   // Support and Events only, the two with a real labour base
  s.addShape('rect', { x: 0.55, y: 3.55, w: 2.90, h: 1.05, fill: { color: C.mint }, line: { type: 'none' } });
  s.addText(`+$${Math.round(gain / 1000)}k a month`, { x: 0.70, y: 3.66, w: 2.6, h: 0.42, isTextBox: true, margin: 0, fontSize: 19, bold: true, color: C.teal, valign: 'middle' });
  s.addText('Five more points of utilisation on Support and Events, at the revenue per chargeable hour each already earns, before any extra cost',
    { x: 0.70, y: 4.06, w: 2.6, h: 0.48, isTextBox: true, margin: 0, fontSize: 8, color: C.body, valign: 'top' });
  note(s, 'Revenue per hour is department revenue divided by chargeable hours. It is not a charge rate, because it carries equipment, hire and freight as well as labour. Cost per hour is direct cost of sales on the same basis. Integration is shown as not applicable: ten chargeable hours against equipment resale produces a ratio, not a rate.',
    { x: 3.60, y: 3.55, w: 3.05, h: 1.05 });
  commentary(s, [
    `Support returns $${money(rows[0].net)} of margin an hour after direct labour. That is the engine.`,
    `Events returns $${money(rows[1].net)} an hour, but on equipment volume, not on a labour rate.`,
    `Integration billed ${money(rows[3].chg)} chargeable hours on ${k(depts[3].revenue)}k of revenue. That work was bought in, not delivered.`,
    `That is why Integration is at ${pct(depts[3].gm_pct)}. We are reselling, not delivering.`,
    `Five points on Support and Events is worth about ${k(gain)}k a month at no extra cost.`,
  ]);
  s.addNotes(`Profit based on utilisation, ${MONTH}.\n\n` + rows.map((r) => `${r.name}: ${pct(r.pctv)} chargeable on ${money(r.chg)} hours. Revenue per chargeable hour $${money(r.revPerHr)}, direct cost per hour $${money(r.labPerHr)}, margin $${money(r.net)}.`).join('\n') + `\n\nTwo cautions on reading this. Revenue per hour is not a charge rate. It is department revenue divided by chargeable hours, so it includes equipment, hire and freight. For Integration that produces a very large number per hour, which simply reflects that the revenue was equipment resale against almost no labour.\n\nThe usable conclusion is the box on the left. Five more points of utilisation on Support and Events, at the revenue per chargeable hour each already earns, is worth roughly ${money(gain)} a month before any additional cost. That is the single largest controllable item in this pack.`);
}

/* ===================================================== 21. section 05 */
section('05', 'Budget and outlook', 'What the rest of the year looks like and what we change',
  'Three months are closed and the remaining nine are still at original budget. That is the decision in front of the board.');

/* ===================================================== 22. revenue by month */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Budget and outlook', 'Revenue by month against budget');
  const m = D.income_by_month;
  const labels = m.map((x) => new Date(x.month).toLocaleString('en-AU', { month: 'short' }));
  s.addChart('bar', [
    { name: 'Actual', labels, values: m.map((x, i) => (i < 3 ? Math.round(x.actual / 1000) : null)) },
    { name: 'Budget', labels, values: m.map((x) => Math.round(x.budget / 1000)) },
  ], {
    x: 0.50, y: 1.58, w: 6.20, h: 2.45, barDir: 'col', barGapWidthPct: 45,
    ...CHART, chartColors: [C.teal, C.grey],
    showLegend: true, legendPos: 'b', legendFontSize: 8, legendColor: C.muted, legendFontFace: '+mn-lt',
    showValue: false,
    valAxisTitle: 'A$ thousands', showValAxisTitle: true, valAxisTitleFontSize: 8, valAxisTitleColor: C.muted, valAxisTitleFontFace: '+mn-lt',
  });
  const q1a = m.slice(0, 3).reduce((a, x) => a + x.actual, 0), q1b = m.slice(0, 3).reduce((a, x) => a + x.budget, 0);
  const restb = m.slice(3).reduce((a, x) => a + x.budget, 0);
  s.addTable([headRow(['A$ thousands', 'Actual', 'Budget', 'Var']),
    row([{ text: 'Q1, Jul to Sep', align: 'left' }, k(q1a), k(q1b), { text: signed((q1a - q1b) / 1000), color: varColour(q1a - q1b) }]),
    row([{ text: 'Oct to Jun, still at budget', align: 'left' }, '-', k(restb), '-']),
    row([{ text: 'Full year', align: 'left' }, k(q1a + restb), k(q1b + restb), { text: signed((q1a - q1b) / 1000), color: varColour(q1a - q1b) }], { bold: true, fill: C.tile }),
  ], { x: 0.55, y: 4.12, w: 6.10, colW: [2.65, 1.15, 1.15, 1.15], border: { type: 'solid', color: C.rule, pt: 0.5 }, autoPage: false });
  commentary(s, [
    `Q1 came in ${k(q1a - q1b)}k behind on revenue, and almost all of it is Consulting and Video.`,
    `The remaining nine months are untouched original budget. They are not a forecast.`,
    `October alone is budgeted at ${k(m[3].budget)}k, nearly double our best month so far.`,
    `December carries a planned loss. The business closes for two weeks, so that is correct and should stay.`,
    `Recommend a Q1 rolling forecast: rebuild Consulting on real capacity and reset the margin assumptions.`,
  ]);
  s.addNotes(`Revenue by month, actual for the three closed months and original budget for the rest.\n\nQ1 actual ${money(q1a)} against a budget of ${money(q1b)}, behind by ${money(q1b - q1a)}.\n\nThe nine months from October are still at the budget set before the year started. Nobody has reforecast them. October is budgeted at ${money(m[3].budget)}, which is higher than any month we have delivered.\n\nTwo things I would say to the board. First, the Q1 variance is concentrated in Consulting and Video, which were budgeted on a resourcing level that does not exist. Second, the margin assumption in the budget is above what the business has delivered for two years, so even a month that hits the revenue number will miss the profit number.\n\nMy recommendation is a rolling forecast from Q2, rebuilding Consulting around actual capacity and resetting gross margin to delivered rates. December should stay as budgeted: the business closes for two weeks and the loss is real.`);
}

/* ===================================================== 23. summary and actions */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Summary and actions', 'What we agreed');
  const acts = [
    ['P&L', `Reforecast FY27 on margin, not revenue. At ${pct(ytdMargin)} gross margin break even needs ${k(breakEven)}k a month and we are running at ${k(revRun)}k.`, 'GC and FC', '31 Oct 26'],
    ['Overhead keys', `Events carries ${k(-depts[1].all_overhead)}k of overhead on ${k(depts[1].revenue)}k of revenue and cannot be profitable on that key. Test the keys for FY28, do not change them mid year.`, 'GC and FC', '30 Nov 26'],
    ['Consulting', `${k(-depts[4].all_overhead)}k of overhead on a ${k(depts[4].revenue)}k fee base. Decide whether the department is rebuilt, merged into Integration, or closed.`, 'GC', '30 Nov 26'],
    ['Customers', `Top three are ${pct(D.top10.slice(0, 3).reduce((a, x) => a + x.ytd, 0) / rev.ytd)} of revenue. Agree a named retention plan, and set the growth target on the second ten.`, 'GC and Sales', '14 Nov 26'],
    ['Cash', `DSO ${D.dso.days.toFixed(1)} days. Getting to ${D.dso.target} releases ${k(D.dso.locked)}k. Chase ${D.debtors[0].name} and ${D.debtors[1].name} first.`, 'FC', '24 Oct 26'],
    ['Timesheets', `Support shows ${money(util[0].non_chargeable)} non-chargeable hours in a month. Correct the coding before utilisation is used to manage anyone.`, 'Ops and FC', '31 Oct 26'],
  ];
  s.addTable([headRow(['Area', 'What we agreed', 'Owner', 'Due'], 'left'),
    ...acts.map(([a, w, o, dd]) => [
      { text: a, options: { fontSize: 8.5, bold: true, color: C.navy, align: 'left', valign: 'top', margin: [4, 5, 4, 5] } },
      { text: w, options: { fontSize: 8.5, color: C.body, align: 'left', valign: 'top', margin: [4, 5, 4, 5] } },
      { text: o, options: { fontSize: 8.5, color: C.body, align: 'left', valign: 'top', margin: [4, 5, 4, 5] } },
      { text: dd, options: { fontSize: 8.5, color: C.body, align: 'left', valign: 'top', margin: [4, 5, 4, 5] } },
    ])], {
    x: 0.55, y: 1.58, w: 8.90, colW: [1.05, 5.70, 1.05, 1.10],
    border: { type: 'solid', color: C.rule, pt: 0.5 }, autoPage: false,
  });
  s.addNotes(`Six actions, each with an owner and a date.\n\nThe first two are the ones that matter. Reforecasting on margin rather than revenue, and testing the overhead keys. Everything else in this pack follows from those two numbers.\n\nThe Consulting decision cannot keep being deferred. It has carried overhead against no fee base for three months.\n\nThe cash action is the only one that produces money this quarter.`);
}

/* ===================================================== 24. thank you */
{
  const s = pres.addSlide({ masterName: 'CTS_DARK' });
  [{ x: 5.5, y: 0.25, d: 2.5, c: C.tealLt }, { x: 6.8, y: 1.35, d: 2.1, c: C.blue }, { x: 5.8, y: 2.35, d: 2.6, c: C.red }]
    .forEach((o, i) => s.addShape('ellipse', { x: o.x, y: o.y, w: o.d, h: o.d, fill: { color: o.c, transparency: 18 }, line: { type: 'none' }, objectName: `closing circle ${i}` }));
  s.addText('cts', { x: 0.55, y: 0.45, w: 2, h: 0.7, isTextBox: true, margin: 0, fontSize: 34, bold: true, color: C.white });
  s.addText('Seamless AV', { x: 0.60, y: 0.98, w: 2, h: 0.25, isTextBox: true, margin: 0, fontSize: 8.5, color: C.grey });
  s.addText('Thank you', { x: 0.55, y: 2.30, w: 5, h: 0.8, isTextBox: true, margin: 0, fontSize: 38, bold: true, color: C.white, valign: 'middle' });
  s.addText('Questions and discussion', { x: 0.55, y: 3.10, w: 5, h: 0.4, isTextBox: true, margin: 0, fontSize: 13, color: C.grey, valign: 'middle' });
  s.addNotes('Close on the two numbers that matter: break even needs 549 thousand a month at our real margin, and we are running at 541. Everything else in the pack is detail underneath that.');
}

module.exports = { pres };
