const B = require('./build.js');
const L = require('./lib.js');
const { C, DEPT } = require('./theme.js');
const { head, commentary, note, tile, row, headRow, CHART, money, k, pct, signed, varColour, G } = L;
const { pres, D, P, rev, gp, oh, eb, np, depts, dTot, byRev, ytdMargin, ohRun, breakEven, revRun, section, MONTH, SHORT } = B;

/* ===================================================== 3. section 01 */
section('01', 'Financial performance', `The P&L for ${MONTH} and the year to date`,
  `Three months of FY27 are closed. Revenue is running at ${k(revRun)} thousand a month and break even needs ${k(breakEven)} thousand at the margin we are actually delivering. That is the whole problem in one line.`);

/* ===================================================== 4. P&L summary */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Financial performance', 'P&L summary');
  const lines = [
    ['Revenue', rev, false], ['Cost of sales', P.cos, false],
    ['Gross profit', gp, true], ['Gross margin %', null, false],
    ['Overheads', oh, false], ['EBITDA', eb, true], ['EBITDA margin %', null, false],
    ['Net profit', np, true],
  ];
  const body = [];
  for (const [label, src, bold] of lines) {
    if (label.endsWith('%')) {
      const base = label[0] === 'G' ? gp : eb;
      const m = base.month / rev.month, mb = base.month_budget / rev.month_budget;
      const y = base.ytd / rev.ytd, yb = base.ytd_budget / rev.ytd_budget;
      const ly = base.month_ly / rev.month_ly;
      body.push(row([
        { text: label, align: 'left' }, pct(m), pct(mb),
        { text: ((m - mb) * 100 >= 0 ? '+' : '') + ((m - mb) * 100).toFixed(1) + ' pts', color: varColour(m - mb) },
        pct(y), pct(yb),
        { text: ((y - yb) * 100 >= 0 ? '+' : '') + ((y - yb) * 100).toFixed(1) + ' pts', color: varColour(y - yb) },
        pct(ly),
      ]));
      continue;
    }
    const mv = src.month - src.month_budget, yv = src.ytd - src.ytd_budget;
    body.push(row([
      { text: label, align: 'left' }, k(src.month), k(src.month_budget),
      { text: signed(mv / 1000), color: varColour(mv) },
      k(src.ytd), k(src.ytd_budget),
      { text: signed(yv / 1000), color: varColour(yv) },
      k(src.month_ly),
    ], { bold, fill: bold ? C.tile : undefined }));
  }
  s.addTable([headRow(['A$ thousands', SHORT, 'Budget', 'Var', 'YTD', 'Budget', 'Var', 'Sep 25']), ...body], {
    x: 0.55, y: 1.60, w: 6.10, colW: [1.55, 0.65, 0.65, 0.65, 0.65, 0.65, 0.65, 0.65],
    border: { type: 'solid', color: C.rule, pt: 0.5 }, autoPage: false,
  });
  note(s, `Revenue beat budget by ${k(rev.month - rev.month_budget)}k and the month still lost ${k(-eb.month)}k of EBITDA. Cost of sales rose ${k(-(P.cos.month - P.cos.month_budget))}k on that revenue, so the gain came in at a margin the plan did not assume.`,
    { x: 0.55, y: 3.98, w: 6.10, h: 0.48 });
  s.addChart('bar', [
    { name: 'Revenue', labels: D.income_by_month.slice(0, 6).map((m) => new Date(m.month).toLocaleString('en-AU', { month: 'short' })), values: D.income_by_month.slice(0, 6).map((m) => Math.round(m.actual / 1000)) },
    { name: 'Budget', labels: D.income_by_month.slice(0, 6).map((m) => new Date(m.month).toLocaleString('en-AU', { month: 'short' })), values: D.income_by_month.slice(0, 6).map((m) => Math.round(m.budget / 1000)) },
  ], {
    x: 0.50, y: 4.32, w: 6.20, h: 0.85, barDir: 'col', barGapWidthPct: 55,
    ...CHART, chartColors: [C.teal, C.grey], showLegend: true, legendPos: 'b',
    legendFontSize: 7.5, legendColor: C.muted, legendFontFace: '+mn-lt',
    valAxisHidden: true, valGridLine: { style: 'none' }, showValue: false,
  });
  commentary(s, [
    `Break even needs ${k(breakEven)}k of revenue a month at our real margin of ${pct(ytdMargin)}. We are running at ${k(revRun)}k.`,
    `${SHORT} beat revenue budget by ${k(rev.month - rev.month_budget)}k and still lost ${k(-eb.month)}k. The miss is margin, not sales.`,
    `Gross margin ${pct(gp.month / rev.month)} against ${pct(gp.month_budget / rev.month_budget)} budget looks good, but last September was ${pct(gp.month_ly / rev.month_ly)}.`,
    `Overheads held at ${k(-oh.month)}k against ${k(-oh.month_budget)}k budget. Cost control is not the issue.`,
    `Year to date EBITDA is ${k(eb.ytd)}k against ${k(eb.ytd_budget)}k planned. Reforecast on margin.`,
  ]);
  s.addNotes(`P&L summary for ${MONTH}.\n\nRevenue of ${money(rev.month)} against a budget of ${money(rev.month_budget)}, ahead by ${money(rev.month - rev.month_budget)} or ${pct((rev.month - rev.month_budget) / rev.month_budget, 0)}. Against last September at ${money(rev.month_ly)} we are up ${pct(rev.month / rev.month_ly - 1, 0)}.\n\nGross profit ${money(gp.month)} at ${pct(gp.month / rev.month)}, against a budget of ${pct(gp.month_budget / rev.month_budget)} and last September at ${pct(gp.month_ly / rev.month_ly)}. The budget margin for the month was unusually low, so beating it is not the achievement it looks like. Year to date margin is ${pct(ytdMargin)} against a plan of ${pct(gp.ytd_budget / rev.ytd_budget)}.\n\nOverheads ${money(-oh.month)} against ${money(-oh.month_budget)} budget, within ${money(Math.abs(oh.month - oh.month_budget))}. The cost base is behaving.\n\nEBITDA ${money(eb.month)} against a budgeted ${money(eb.month_budget)}. Net profit ${money(np.month)} after depreciation and amortisation of ${money(-P.da.month)}.\n\nYear to date: revenue ${money(rev.ytd)} against ${money(rev.ytd_budget)}, behind by ${money(rev.ytd_budget - rev.ytd)}. EBITDA ${money(eb.ytd)} against ${money(eb.ytd_budget)}.\n\nThe point to make in the room: at ${pct(ytdMargin)} gross margin we need ${money(breakEven)} of revenue every month to cover the overhead base. We are averaging ${money(revRun)}. The gap is structural and it will not close by selling more at this margin.`);
}

/* ===================================================== 5. P&L per department */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Financial performance', `P&L per department, ${SHORT}`);
  const body = depts.map((x) => row([
    { text: x.name, align: 'left' }, k(x.revenue), k(x.budget),
    { text: signed((x.revenue - x.budget) / 1000), color: varColour(x.revenue - x.budget) },
    k(x.gross_profit), pct(x.gm_pct), k(x.all_overhead),
    { text: k(x.result), color: varColour(x.result) },
  ]));
  body.push(row([
    { text: 'CTS', align: 'left' }, k(dTot.revenue), k(dTot.budget),
    { text: signed((dTot.revenue - dTot.budget) / 1000), color: varColour(dTot.revenue - dTot.budget) },
    k(dTot.gross_profit), pct(dTot.gross_profit / dTot.revenue), k(dTot.all_overhead),
    { text: k(dTot.result), color: varColour(dTot.result) },
  ], { bold: true, fill: C.tile }));
  s.addTable([headRow(['A$ thousands', 'Revenue', 'Budget', 'Var', 'Gross profit', 'GM %', 'All overhead', 'Result']), ...body], {
    x: 0.55, y: 1.60, w: 6.10, colW: [1.25, 0.72, 0.65, 0.58, 0.82, 0.6, 0.78, 0.7],
    border: { type: 'solid', color: C.rule, pt: 0.5 }, autoPage: false,
  });
  note(s, "All overhead is the department's own overhead plus its share of admin, spread on the keys set in the pack. Admin itself carries nothing once its overhead is charged out. Result is revenue less direct costs less all overhead, and the five departments add back to net profit exactly.",
    { x: 0.55, y: 3.50, w: 6.10, h: 0.7 });
  commentary(s, [
    `Only Support covers its overhead. It returned ${k(depts[0].result)}k on ${k(depts[0].revenue)}k of revenue.`,
    `Events carries ${k(-depts[1].all_overhead)}k of overhead on ${k(depts[1].revenue)}k of revenue, so ${pct(depts[1].gm_pct)} margin still loses ${k(-depts[1].result)}k.`,
    `Integration booked ${k(depts[3].revenue)}k against ${k(depts[3].budget)}k budget, but at ${pct(depts[3].gm_pct)} margin returned only ${k(depts[3].result)}k.`,
    `Consulting holds ${k(-depts[4].all_overhead)}k of overhead on a ${k(depts[4].revenue)}k fee base. That needs a decision, not a forecast.`,
    `The overhead keys suit a different shape of business. Retest before FY28.`,
  ]);
  s.addNotes(`Every department against budget for ${MONTH}, with the full overhead charge.\n\n` + depts.map((x) => `${x.name}: revenue ${money(x.revenue)} against ${money(x.budget)} budget, ${x.revenue >= x.budget ? 'ahead' : 'behind'} by ${money(Math.abs(x.revenue - x.budget))}. Gross profit ${money(x.gross_profit)} at ${pct(x.gm_pct)}. Overhead ${money(-x.all_overhead)}, of which ${money(-x.alloc_overhead)} is allocated admin. Result ${money(x.result)}.`).join('\n\n') + `\n\nThe five departments total ${money(dTot.result)}, which is net profit for the month.\n\nTwo things to flag. First, Integration's ${money(depts[3].revenue - depts[3].budget)} beat on revenue converted to almost nothing because the margin was ${pct(depts[3].gm_pct)} on equipment resale. Second, Events absorbs ${money(-depts[1].alloc_overhead)} of allocated admin on ${money(depts[1].revenue)} of revenue. The key charging it is the premises key, and Events is not the department consuming the premises.`);
}

/* ===================================================== 6. three months to date */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Financial performance', 'The three months to date');
  const mk = (label, b, c, d, ytd, ytdb, bold) => row([
    { text: label, align: 'left' }, k(b), k(c), k(d), k(ytd), k(ytdb),
    { text: signed((ytd - ytdb) / 1000), color: varColour(ytd - ytdb) },
    { text: ytdb ? pct((ytd - ytdb) / Math.abs(ytdb), 0) : '-', color: varColour(ytd - ytdb) },
  ], { bold, fill: bold ? C.tile : undefined });
  const S = D.months || null;
  const body = [
    mk('Total income', 459048.71, 603466.54, rev.month, rev.ytd, rev.ytd_budget, true),
    ...byRev.map((x) => row([
      { text: '    ' + x.name, align: 'left' }, '', '', k(x.revenue), k(x.ytd_revenue), k(x.ytd_budget),
      { text: signed((x.ytd_revenue - x.ytd_budget) / 1000), color: varColour(x.ytd_revenue - x.ytd_budget) },
      { text: pct((x.ytd_revenue - x.ytd_budget) / Math.abs(x.ytd_budget), 0), color: varColour(x.ytd_revenue - x.ytd_budget) },
    ])),
    mk('Cost of sales', -282514.16, -335516.00, P.cos.month, P.cos.ytd, P.cos.ytd_budget, true),
    mk('Gross profit', 176534.55, 267950.54, gp.month, gp.ytd, gp.ytd_budget, true),
    mk('Overheads', -202397.94, -208876.19, oh.month, oh.ytd, oh.ytd_budget, true),
    mk('EBITDA', -25863.39, 59074.35, eb.month, eb.ytd, eb.ytd_budget, true),
  ];
  s.addTable([headRow(['A$ thousands', 'Jul 26', 'Aug 26', 'Sep 26', 'YTD', 'Budget', 'Var', 'Var %']), ...body], {
    x: 0.55, y: 1.55, w: 6.10, colW: [1.42, 0.67, 0.67, 0.67, 0.67, 0.67, 0.67, 0.66],
    border: { type: 'solid', color: C.rule, pt: 0.5 }, autoPage: false,
  });
  note(s, 'Revenue lines are shown by department for the month and the year to date. The quarter reads as one trading period, which is how the budget was built.',
    { x: 0.55, y: 4.70, w: 6.10, h: 0.45 });
  commentary(s, [
    `Three closed months: ${k(459048.71)}k, ${k(603466.54)}k and ${k(rev.month)}k. A quarter of ${k(rev.ytd)}k against ${k(rev.ytd_budget)}k planned.`,
    `The whole ${k(rev.ytd_budget - rev.ytd)}k shortfall is Consulting and Video. Support and Events are on plan.`,
    `Gross profit fell from ${pct(267950.54 / 603466.54)} in August to ${pct(gp.month / rev.month)} in September on similar revenue.`,
    `Overheads are flat across the quarter at about ${k(ohRun)}k a month. That is the number revenue has to clear.`,
    `A quarter of trading has produced ${k(eb.ytd)}k of EBITDA. The plan said ${k(eb.ytd_budget)}k.`,
  ]);
  s.addNotes(`The quarter in one view.\n\nJuly ${money(459048.71)}, August ${money(603466.54)}, September ${money(rev.month)}. Year to date ${money(rev.ytd)} against a budget of ${money(rev.ytd_budget)}, behind by ${money(rev.ytd_budget - rev.ytd)}.\n\nBy department year to date: ` + byRev.map((x) => `${x.name} ${money(x.ytd_revenue)} against ${money(x.ytd_budget)}`).join(', ') + `.\n\nGross profit for the quarter ${money(gp.ytd)} at ${pct(ytdMargin)} against a planned ${pct(gp.ytd_budget / rev.ytd_budget)}. Overheads ${money(-oh.ytd)} against ${money(-oh.ytd_budget)}, favourable by ${money(oh.ytd - oh.ytd_budget)}.\n\nEBITDA ${money(eb.ytd)} against ${money(eb.ytd_budget)}. The business has traded a quarter and produced roughly break even.`);
}
