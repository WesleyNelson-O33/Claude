const B = require('./build.js');
const L = require('./lib.js');
const { C, DEPT } = require('./theme.js');
const { head, commentary, note, tile, row, headRow, CHART, money, k, pct, signed, varColour } = L;
const { pres, D, P, rev, gp, oh, eb, np, depts, dTot, byRev, ytdMargin, ohRun, breakEven, revRun, section, MONTH, SHORT } = B;

const names = byRev.map((d) => d.name);
const colours = names.map((n) => DEPT[n]);

/* ===================================================== 7. revenue vs budget and last year */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Financial performance', 'Revenue against budget and last year');
  s.addChart('bar', [
    { name: `${SHORT} actual`, labels: names, values: byRev.map((d) => Math.round(d.revenue / 1000)) },
    { name: `${SHORT} budget`, labels: names, values: byRev.map((d) => Math.round(d.budget / 1000)) },
    { name: 'Sep 25 actual', labels: names, values: byRev.map((d) => Math.round(d.ly / 1000)) },
  ], {
    x: 0.50, y: 1.60, w: 6.20, h: 3.35, barDir: 'col', barGapWidthPct: 45,
    ...CHART, chartColors: [C.teal, C.grey, C.tealLt],
    showLegend: true, legendPos: 'b', legendFontSize: 8, legendColor: C.muted, legendFontFace: '+mn-lt',
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 7.5,
    valAxisTitle: 'A$ thousands', showValAxisTitle: true, valAxisTitleFontSize: 8, valAxisTitleColor: C.muted, valAxisTitleFontFace: '+mn-lt',
  });
  commentary(s, [
    `Integration booked ${k(depts[3].revenue)}k against a ${k(depts[3].budget)}k budget. One fit-out, not a run rate.`,
    `Support delivered ${k(depts[0].revenue)}k against ${k(depts[0].budget)}k. On plan, as contract income should be.`,
    `Events at ${k(depts[1].revenue)}k beat a ${k(depts[1].budget)}k budget, but September is a quiet month either side of results season.`,
    `Video ${k(depts[2].revenue)}k against ${k(depts[2].budget)}k and Consulting ${k(depts[4].revenue)}k against ${k(depts[4].budget)}k. Both are structurally short.`,
    `Company revenue is ${pct(rev.month / rev.month_ly - 1, 0)} up on last September. Growth is real; the margin on it is not.`,
  ]);
  s.addNotes(`Revenue by department, ${MONTH} actual against budget and against last September.\n\n` + byRev.map((d) => `${d.name}: ${money(d.revenue)} against a budget of ${money(d.budget)} and ${money(d.ly)} last year.`).join('\n') + `\n\nCompany revenue ${money(rev.month)} against ${money(rev.month_budget)} budget and ${money(rev.month_ly)} last September.\n\nThe beat is Integration. It is one large fit-out at ${pct(depts[3].gm_pct)} gross margin, so it adds revenue and very little profit. Read it as a timing event, not as growth. Video and Consulting are both short of budget again, which is the third month in a row.`);
}

/* ===================================================== 8. contribution */
{
  const { barBlock } = require('./part0.js');
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Financial performance', 'Where the profit is made and lost');
  const res = byRev.map((d) => ({ name: d.name, value: Math.round(d.result / 1000), colour: DEPT[d.name] }));
  const gms = byRev.map((d) => ({ name: d.name, value: +(d.gm_pct * 100).toFixed(1), colour: DEPT[d.name] }));
  s.addText('Result after all overhead, A$ thousands', { x: 0.55, y: 1.54, w: 6.1, h: 0.22, isTextBox: true, margin: 0, fontSize: 8.5, bold: true, color: C.muted });
  barBlock(s, { x: 0.55, y: 1.80, w: 6.10, rowH: 0.29, rows: res, min: Math.min(0, ...res.map((r) => r.value)) * 1.08, max: Math.max(...res.map((r) => r.value)) * 1.25, fmt: (v) => (v >= 0 ? '' : '(') + Math.abs(v) + (v >= 0 ? '' : ')') });
  s.addText('Gross margin %, before any overhead', { x: 0.55, y: 3.32, w: 6.1, h: 0.22, isTextBox: true, margin: 0, fontSize: 8.5, bold: true, color: C.muted });
  barBlock(s, { x: 0.55, y: 3.58, w: 6.10, rowH: 0.29, rows: gms, min: 0, max: Math.max(...gms.map((r) => r.value)) * 1.18, fmt: (v) => v.toFixed(1) + '%' });
  commentary(s, [
    'Two of five departments cover their overhead. Three do not.',
    `Margin and result do not line up. Events earns ${pct(depts[1].gm_pct)} and still loses ${k(-depts[1].result)}k.`,
    `Integration earns ${pct(depts[3].gm_pct)} and returns ${k(depts[3].result)}k, because it carries almost no overhead.`,
    'So the overhead keys, not the trading, decide who looks profitable.',
    'Fix the keys before anyone is judged on these lines.',
  ]);
  s.addNotes(`Two blocks, deliberately stacked. The top is the result after all overhead. The bottom is gross margin before any overhead.\n\nRead them against each other. Events has the second highest gross margin in the business at ${pct(depts[1].gm_pct)} and the worst result at ${money(depts[1].result)}. Integration has the lowest margin at ${pct(depts[3].gm_pct)} and a positive result of ${money(depts[3].result)}.\n\nThe difference is the overhead charge. Events carries ${money(-depts[1].all_overhead)} and Integration carries ${money(-depts[3].all_overhead)}.\n\nThat is an allocation outcome, not a trading outcome. The premises key puts most of the admin cost onto Events, and Events is not the department consuming the premises. Until that key is tested these results should not be used to judge the departments.`);
}

/* ===================================================== 9. section 02 */
section('02', 'Cost base', 'Where the money goes and which lines move',
  `Overheads are steady at about ${k(ohRun)} thousand a month. The cost base is not the problem this month, but it is the number revenue has to clear, so it is worth seeing in full.`);

/* ===================================================== 10. operating cost distribution */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Cost base', 'Operating cost distribution, year to date');
  const cd = D.cost_dist.filter((x) => Math.abs(x.amount) > 500)
    .map((x) => ({ ...x, label: x.label.replace('Cost of sales, ', 'COS, ').replace('Overhead, ', '').replace(/ Exp$/, '') }))
    .sort((a, b) => Math.abs(b.amount) - Math.abs(a.amount));
  s.addChart('bar', [{ name: 'Cost', labels: cd.map((x) => x.label), values: cd.map((x) => Math.round(Math.abs(x.amount) / 1000)) }], {
    x: 0.50, y: 1.60, w: 6.20, h: 3.30, barDir: 'bar', barGapWidthPct: 40,
    ...CHART, chartColors: [C.teal, C.teal, C.tealLt, C.tealLt, C.grey, C.grey, C.grey, C.grey],
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 8.5,
    catAxisOrientation: 'maxMin', valAxisHidden: true, valGridLine: { style: 'none' },
  });
  note(s, 'A$ thousands, year to date. Cost of sales and overhead on one scale, so the size of each line is comparable.',
    { x: 0.55, y: 4.92, w: 6.10, h: 0.3 });
  const empl = D.cost_dist.find((x) => x.label.includes('Employment'));
  const lab = D.cost_dist.find((x) => x.label.includes('labour'));
  commentary(s, [
    `Direct labour ${k(-lab.amount)}k and employment overhead ${k(-empl.amount)}k together are ${pct((Math.abs(lab.amount) + Math.abs(empl.amount)) / (rev.ytd - eb.ytd), 0)} of all cost.`,
    `People are the cost base. Everything else is rounding by comparison.`,
    `Employment overhead at ${k(-empl.amount)}k is fixed within the year. It does not flex when a month comes in short.`,
    `Equipment and services ${k(-D.cost_dist.find((x) => x.label.includes('equipment')).amount)}k is the only genuinely variable block.`,
    `Cost reduction means headcount or premises. There is no third lever of any size.`,
  ]);
  s.addNotes(`Operating cost distribution for the year to date, cost of sales and overhead together.\n\n` + cd.map((x) => `${x.label} ${money(Math.abs(x.amount))} at ${pct(x.share)}`).join('. ') + `.\n\nThe message is concentration. Direct labour and employment overhead together are the clear majority of the cost base. Employment overhead is fixed inside the year, which is why a short revenue month drops straight to the bottom line.\n\nWhen the board asks what we can cut, the honest answer is that the only lines big enough to matter are people and premises. Trimming subscriptions and travel will not change the shape of this.`);
}

/* ===================================================== 11. other operating expenses */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Cost base', 'Other operating expenses, year to date');
  const oo = D.other_opex.filter((x) => Math.abs(x.amount) > 100).sort((a, b) => Math.abs(b.amount) - Math.abs(a.amount));
  s.addChart('bar', [{ name: 'Cost', labels: oo.map((x) => x.label), values: oo.map((x) => Math.round(Math.abs(x.amount) / 1000)) }], {
    x: 0.50, y: 1.60, w: 6.20, h: 3.20, barDir: 'bar', barGapWidthPct: 40,
    ...CHART, chartColors: [C.teal, C.teal, C.tealLt, C.grey, C.grey, C.grey, C.grey, C.grey, C.grey],
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 8.5,
    catAxisOrientation: 'maxMin', valAxisHidden: true, valGridLine: { style: 'none' },
  });
  const rent = oo.find((x) => x.label === 'Rent'), outg = oo.find((x) => x.label.includes('Outgoings')), it = oo.find((x) => x.label.includes('IT Network'));
  note(s, `A$ thousands, year to date. Operating expenses in total are ${k(Math.abs(D.cost_dist.find((x) => x.label.includes('Operating Expenses')).amount))}k.`,
    { x: 0.55, y: 4.85, w: 6.10, h: 0.3 });
  commentary(s, [
    `Premises are ${pct((Math.abs(rent.amount) + Math.abs(outg.amount)) / Math.abs(D.cost_dist.find((x) => x.label.includes('Operating Expenses')).amount), 0)} of this block: rent ${k(-rent.amount)}k plus outgoings ${k(-outg.amount)}k.`,
    `IT network at ${k(-it.amount)}k is the second line and the only one with real variability.`,
    `Everything below the top three is under ${k(6000)}k across a full quarter.`,
    `There is no saving of consequence in the tail. Do not spend management time there.`,
    `Premises is the lever, and it is a lease decision, not a budget decision.`,
  ]);
  s.addNotes(`Other operating expenses, year to date.\n\n` + oo.map((x) => `${x.label} ${money(Math.abs(x.amount))} at ${pct(x.share)}`).join('. ') + `.\n\nRent and rental outgoings together are the dominant cost and they are fixed by lease. IT network service and support is the second line.\n\nThe practical conclusion is that this block is not where cost comes out. If the business needs to reduce operating expense materially, it is a premises conversation.`);
}

/* ===================================================== 12. labour mix */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Cost base', 'Labour mix, year to date');
  const lm = D.labour_mix.filter((x) => Math.abs(x.amount) > 500 && !/^all |total|check|in total/i.test(x.label))
    .sort((a, b) => Math.abs(b.amount) - Math.abs(a.amount)).slice(0, 8);
  s.addChart('bar', [{ name: 'Cost', labels: lm.map((x) => x.label), values: lm.map((x) => Math.round(Math.abs(x.amount) / 1000)) }], {
    x: 0.50, y: 1.60, w: 6.20, h: 3.20, barDir: 'bar', barGapWidthPct: 40,
    ...CHART, chartColors: [C.teal, C.teal, C.tealLt, C.tealLt, C.grey, C.grey, C.grey, C.grey],
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 8.5,
    catAxisOrientation: 'maxMin', valAxisHidden: true, valGridLine: { style: 'none' },
  });
  note(s, 'A$ thousands, year to date. Direct labour charged to cost of sales, split by type.',
    { x: 0.55, y: 4.85, w: 6.10, h: 0.3 });
  const tot = lm.reduce((a, x) => a + Math.abs(x.amount), 0);
  commentary(s, [
    `Direct labour is ${k(tot)}k for the quarter across ${lm.length} lines.`,
    `${lm[0].label} at ${k(Math.abs(lm[0].amount))}k is ${pct(Math.abs(lm[0].amount) / tot, 0)} of it.`,
    `Sub contract labour is the flexible part. Everything else is payroll.`,
    `Wages and salaries move with hours worked, not with hours billed, which is the utilisation point later in this pack.`,
    `Watch the sub contract line: it rises when the permanent team runs out of capacity.`,
  ]);
  s.addNotes(`Labour mix for the year to date.\n\n` + lm.map((x) => `${x.label} ${money(Math.abs(x.amount))}`).join('. ') + `.\n\nThe great majority is permanent payroll. Sub contract labour is the only line that flexes with volume.\n\nThis is the link to the utilisation section. We pay for hours worked. We bill for hours that are chargeable. Where those two diverge, margin goes with it.`);
}

/* ===================================================== 13. employment split */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Cost base', 'Employment expense, office and departmental');
  const st = D.staff, fte = D.fte;
  const byName = Object.fromEntries(fte.map((f) => [f.name, f]));
  const body = st.map((x) => row([
    { text: x.name, align: 'left' }, money(x.departmental), money(x.office), money(x.total),
    (byName[x.name].hours / D.annual_hours).toFixed(2), pct(byName[x.name].share),
  ]));
  const tot = st.reduce((a, x) => ({ d: a.d + x.departmental, o: a.o + x.office, t: a.t + x.total }), { d: 0, o: 0, t: 0 });
  body.push(row([{ text: 'Total', align: 'left' }, money(tot.d), money(tot.o), money(tot.t),
    (fte.reduce((a, f) => a + f.hours, 0) / D.annual_hours).toFixed(2), '100.0%'], { bold: true, fill: C.tile }));
  s.addTable([headRow(['September 2026', 'Departmental staff', 'Office staff', 'Total', 'FTE', 'Share']), ...body], {
    x: 0.55, y: 1.60, w: 6.10, colW: [1.25, 1.22, 1.05, 1.0, 0.75, 0.83],
    border: { type: 'solid', color: C.rule, pt: 0.5 }, autoPage: false,
  });
  note(s, `Departmental staff is employment expense coded to the department's own cost centre. Office staff is the admin employment pot of ${money(tot.o)}, charged out on each department's FTE. FTE is rolling twelve month hours on that cost centre divided by ${money(D.annual_hours)} working hours a year. The split now calculates in the pack; nothing here is typed.`,
    { x: 0.55, y: 3.45, w: 6.10, h: 0.9 });
  commentary(s, [
    `The admin employment pot is ${money(tot.o)} for the month and it is charged out in full.`,
    `Support carries ${pct(st[0].office / tot.o)} of it on ${byName.Support.hours.toLocaleString('en-AU', { maximumFractionDigits: 0 })} rolling hours.`,
    `Events pays more for its own staff, ${money(st[1].departmental)}, than it does for office staff.`,
    `Integration takes almost no office cost because almost nothing is clocked to it.`,
    `This is the line that used to be typed by hand. It is now driven by the labour hours report.`,
  ]);
  s.addNotes(`Employment expense split between office staff and departmental staff, ${MONTH}.\n\nDepartmental staff is the employment expense sitting on each department's own cost centre: ` + st.map((x) => `${x.name} ${money(x.departmental)}`).join(', ') + `.\n\nOffice staff is the admin employment pot of ${money(tot.o)}, charged across the departments on their share of full time equivalent: ` + st.map((x) => `${x.name} ${money(x.office)} at ${pct(byName[x.name].share)}`).join(', ') + `.\n\nFTE is rolling twelve month worked hours on that cost centre divided by ${money(D.annual_hours)} hours, which is 253 working days at eight hours. That is the method in the training recording.\n\nThe important change this month is that this no longer relies on a typed figure. The pack calculates it from the Labour Hours Report, so the split reproduces every month without anyone deciding it.`);
}
