const B = require('./build.js');
const L = require('./lib.js');
const { C, DEPT } = require('./theme.js');
const { head, commentary, headRow, row, money, k, pct, signed, varColour } = L;
const { placeSnip, source } = require('./snips.js');
const { EB } = require('./part0.js');
const { pres, D, P, rev, gp, oh, eb, np, depts, dTot, byRev, util, utilTot, staff,
        ytdMargin, ohRun, breakEven, revRun, section, MONTH, SHORT,
        LY_NAME, MONTH_END, DONE, LEFT, HER, HU, HOURS, SHOW_PEOPLE } = B;
const { YR } = require('./part0.js');
const PRIOR = eb.ytd - eb.month;      // the months of this year before this one

const BOX = { x: 0.55, y: 1.58, w: 6.10, h: 3.38 };
const FROM = (tab, range) => `Snipped from the Financial Controller Pack, ${tab} tab, ${range}.`;

// A snip wider than about 3:1 is short on a 6.1 inch column, so those run the
// full content width with the commentary as a strip underneath instead.
const WIDE = 3.0;

/** A slide that is one Excel snip plus the controller's read of it. */
function snipSlide(o) {
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, o.eyebrow, o.title);
  const [pw, ph] = require('./snips.js').SIZES[o.snip];
  const wide = pw / ph >= WIDE;
  const box = wide ? { x: 0.55, y: 1.50, w: 8.90, h: 2.18 } : { ...BOX, ...(o.box || {}) };
  const r = placeSnip(s, o.snip, box, wide);
  source(s, o.source, wide ? Math.min(r.bottom + 0.06, 3.74) : r.bottom + 0.08, box.w, box.x);
  if (wide) {
    const PY = 4.02, PH = 1.04;                 // fixed strip, clear of the footer
    const bullets = o.commentary.slice(0, 4);   // four fit two up, five do not
    const half = Math.ceil(bullets.length / 2);
    // Size the strip to the longer of the two columns, same as the side panel.
    const longer = bullets.slice(0, half).join('').length >=
                   bullets.slice(half).join('').length
                 ? bullets.slice(0, half) : bullets.slice(half);
    const strip = L.fitPanel(longer, 4.05, PH - 0.36, [8.5, 8, 7.5, 7, 6.5]);
    s.addShape('rect', { x: 0.55, y: PY, w: 8.90, h: PH, fill: { color: C.mint },
      line: { type: 'none' }, objectName: 'commentary panel' });
    s.addText('COMMENTARY', { x: 0.75, y: PY + 0.09, w: 3, h: 0.18, isTextBox: true,
      margin: 0, fontSize: 8.5, bold: true, color: C.teal, charSpacing: 1.8 });
    [bullets.slice(0, half), bullets.slice(half)].forEach((col, i) => {
      s.addText(col.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < col.length - 1 } })), {
        x: 0.75 + i * 4.35, y: PY + 0.29, w: 4.05, h: PH - 0.36,
        isTextBox: true, margin: 0, ...strip, color: C.body,
        paraSpaceAfter: 3, valign: 'top', shrinkText: true,
      });
    });
  } else {
    commentary(s, o.commentary);
  }
  if (o.notes) s.addNotes(o.notes);
  return s;
}

/* ===================================================== 01 financial performance */
section('01', 'Financial performance', `The P&L for ${MONTH} and the year to date`,
  `${DONE} months of FY27 are closed. Revenue is running at ${k(revRun)} thousand a month and break even needs ${k(breakEven)} thousand at the margin we actually deliver. ${revRun >= breakEven ? 'We are above that line for the first time this year.' : 'We are still below that line.'}`);

snipSlide({
  eyebrow: 'Financial performance', title: 'P&L summary, income',
  snip: 'pl_income', source: FROM('Summary', 'income by department, the last three months and year to date'),
  commentary: [
    `${SHORT} sold ${k(rev.month)}k against a ${k(rev.month_budget)}k plan. That is ${k(rev.month - rev.month_budget)}k more than planned and our biggest month of the year.`,
    `We kept ${pct(gp.month / rev.month, 0)} of every dollar sold. The plan said ${pct(gp.month_budget / rev.month_budget, 0)} and last ${LY_NAME} we kept ${pct(gp.month_ly / rev.month_ly, 0)}.`,
    `So gross profit came in at ${k(gp.month)}k, which is ${k(Math.abs(gp.month - gp.month_budget))}k under plan even though sales beat plan.`,
    `Running costs were ${k(-oh.month)}k against ${k(-oh.month_budget)}k planned. Spending is steady and is not what moved the month.`,
    `We need ${k(breakEven)}k of sales a month to cover our costs at the margin we deliver. Year to date we are averaging ${k(revRun)}k.`,
  ],
  notes: `P&L summary, straight off the Summary tab with the outline shut, so this is the view on screen.\n\nRevenue ${money(rev.month)} against a budget of ${money(rev.month_budget)}, ahead by ${money(rev.month - rev.month_budget)}. Against last ${LY_NAME} at ${money(rev.month_ly)} we are up ${pct(rev.month / rev.month_ly - 1, 0)}. It is the largest month of the financial year so far.\n\nGross profit ${money(gp.month)} at ${pct(gp.month / rev.month)}, against a budget of ${pct(gp.month_budget / rev.month_budget)} and last ${LY_NAME} at ${pct(gp.month_ly / rev.month_ly)}. Sales beat budget by ${money(rev.month - rev.month_budget)} and gross profit still came in ${money(Math.abs(gp.month - gp.month_budget))} under it. That is the whole month in one sentence.\n\nOverheads ${money(-oh.month)} against ${money(-oh.month_budget)} budget, within ${money(Math.abs(oh.month - oh.month_budget))}. The cost base is behaving.\n\nEBITDA ${money(eb.month)} against a budgeted ${money(eb.month_budget)}. Net profit ${money(np.month)} after the instant asset write off of ${money(-P.da.month)}.\n\nYear to date: revenue ${money(rev.ytd)} against ${money(rev.ytd_budget)}, which is ${pct(rev.ytd / rev.ytd_budget, 0)} of plan. Gross profit ${money(gp.ytd)} against ${money(gp.ytd_budget)}, which is only ${pct(gp.ytd / gp.ytd_budget, 0)} of plan. The gap between those two percentages is the margin problem, and it has not changed.\n\nEBITDA year to date ${money(eb.ytd)} against ${money(eb.ytd_budget)}.\n\nThe line to leave in the room: at ${pct(ytdMargin)} gross margin we need ${money(breakEven)} of revenue every month to cover overhead. We are averaging ${money(revRun)}. ${revRun >= breakEven ? 'We are now above that line, on volume rather than on margin. Hold the volume and the year works. Lose it and the margin cannot carry us.' : 'We are still under that line.'}`,
});

snipSlide({
  eyebrow: 'Financial performance', title: 'P&L summary, cost and profit',
  snip: 'pl_result', source: FROM('Summary', 'cost of sales, overheads and the profit lines'),
  commentary: [
    `Operating profit, or EBITDA, is ${k(YR.ytd_actual)}k for the ${DONE} months to date against ${k(YR.ytd_budget)}k planned. We are ${k(-YR.shortfall)}k short of plan.`,
    `${SHORT} on its own made ${k(eb.month)}k of it. The ${DONE - 1} months before it made ${k(PRIOR)}k between them.`,
    `To still hit the ${k(YR.fy_budget)}k full year target the ${LEFT} months left have to make ${k(YR.still_needed)}k, which is ${k(YR.per_month_needed)}k a month.`,
    `${k(YR.per_month_needed)}k a month is below what ${SHORT} alone delivered. On this month's shape the full year budget is back in reach.`,
    `The problem is still what we keep, not what we sell. Sales hit ${pct(rev.ytd / rev.ytd_budget, 0)} of plan. What was left after direct costs hit only ${pct(gp.ytd / gp.ytd_budget, 0)}.`,
  ],
  notes: `Where the year stands on profit, after ${DONE} months.\n\nEBITDA year to date is ${money(YR.ytd_actual)} against a budget of ${money(YR.ytd_budget)}. That is ${money(-YR.shortfall)} short of plan.\n\nHow we got here matters more than the total. The ${DONE - 1} months before ${MONTH} produced ${money(PRIOR)} between them. ${MONTH} on its own produced ${money(eb.month)}, which is more than those ${DONE - 1} months together and more than the whole of last financial year's ${money(YR.fy26_actual)}.\n\nWhat that does to the full year. The FY27 budget is ${money(YR.fy_budget)} of EBITDA. We have ${money(YR.ytd_actual)} of it after ${DONE} months, so the ${LEFT} months left have to produce ${money(YR.still_needed)}, or ${money(YR.per_month_needed)} a month. ${MONTH} delivered ${money(eb.month)}. On that basis the full year budget is reachable, which is a different position from the one I reported at the end of the quarter.\n\nThe caution against it. One month is not a trend, and this one is carried by volume rather than by margin. Revenue year to date is ${pct(rev.ytd / rev.ytd_budget, 1)} of budget, which is respectable. Gross profit is only ${pct(gp.ytd / gp.ytd_budget, 1)} of budget. We are selling close to plan and keeping well under it, and that has been true every month this year.\n\nSpending is not the cause. Overheads ran at ${money(-oh.ytd)} for the ${DONE} months, ${money(ohRun)} a month, and came in ${money(Math.abs(oh.ytd - oh.ytd_budget))} ${oh.ytd > oh.ytd_budget ? 'under' : 'over'} budget. Wages and the office are what we spend and both are locked in for the year, so when we keep less of each sale it goes straight off the profit.\n\nWhere that leaves break even. At the ${pct(ytdMargin)} gross margin we actually deliver, the business needs ${money(breakEven)} of revenue a month to cover overhead. We are running at ${money(revRun)}. ${revRun >= breakEven ? 'We are above it, and ' + MONTH + ' is the month that put us there.' : 'We are still under it.'}\n\nTwo things still to watch. December is budgeted at a loss because the business closes for two weeks, so Q2 will not look like this month. And ${money(D.dso.locked)} of cash is sitting in debtors above the ${D.dso.target} day target, which is more than it was at the end of the quarter.\n\nWhat I am recommending. Reforecast on margin rather than revenue, because the margin gap is the one thing that has not improved. Reset the gross margin assumption to the rate delivered over two years rather than the rate in the plan. Hold December as budgeted. And do not revise the full year EBITDA down on the quarter alone; ${MONTH} has changed that picture and the half year is the right point to decide.`,
});

snipSlide({
  eyebrow: 'Financial performance', title: `Department result, ${SHORT}`,
  snip: 'dept_result', source: FROM('Overhead Split', 'section 7, department result'),
  commentary: (() => {
    const win = [...depts].sort((a, b) => b.result - a.result);
    const lose = depts.filter((x) => x.result < 0);
    return [
      `${win[0].name} made ${k(win[0].result)}k on ${k(win[0].revenue)}k of sales, after its ${k(-win[0].all_overhead)}k share of running costs. The best month any department has had this year.`,
      `${win[1].name} made ${k(win[1].result)}k on ${k(win[1].revenue)}k, keeping ${pct(win[1].gm_pct, 0)} of what it sold.`,
      `Integration billed ${k(depts[3].revenue)}k against a ${k(depts[3].budget)}k plan but keeps only ${pct(depts[3].gm_pct, 0)} of it, so it made ${k(depts[3].result)}k.`,
      lose.length
        ? `${lose.map((x) => x.name).join(' and ')} ${lose.length > 1 ? 'are' : 'is'} the only ${lose.length > 1 ? 'ones' : 'one'} under water, by ${k(-lose.reduce((a, x) => a + x.result, 0))}k between ${lose.length > 1 ? 'them' : 'it'}.`
        : `Every department covered its share of running costs this month.`,
      `The five departments add up to the company result exactly. The check row on the tab proves it.`,
    ];
  })(),
  notes: `Every department against budget with the full overhead charge, read off section 7 of the Overhead Split tab.\n\n` +
    depts.map((x) => `${x.name}: revenue ${money(x.revenue)} against ${money(x.budget)} budget. Gross profit ${money(x.gross_profit)} at ${pct(x.gm_pct)}. Overhead ${money(-x.all_overhead)}, of which ${money(-x.alloc_overhead)} is allocated admin. Result ${money(x.result)}.`).join('\n') +
    `\n\nThe five departments total ${money(dTot.result)}, which is the operating result for the month. Net profit of ${money(np.month)} is that plus other income of ${money(P.other_income.month)}.\n\nTwo things to flag. Integration sold ${money(depts[3].revenue - depts[3].budget)} more than planned and it turned into very little profit, because we keep only ${pct(depts[3].gm_pct)} on reselling equipment. And ${depts[0].name} billed ${money(Math.abs(depts[0].revenue - depts[0].budget))} ${depts[0].revenue < depts[0].budget ? 'less' : 'more'} than plan at a ${pct(depts[0].gm_pct)} margin, which is why the department that normally carries the company is the one in the red this month. That is worth asking the question about before next month.`,
});

/* ------------------------- the three departments the P&L is built on ------ */
// Her Summary tab has three trading departments, not five: Production is the
// parent of Events and Video, Consulting is the parent of Integration and
// Consulting. Section 8 of Overhead Split rolls section 7 up to those three.
const roll = (name, list) => {
  const t = (f) => list.reduce((a, d) => a + d[f], 0);
  const o = { name, revenue: t('revenue'), budget: t('budget'),
              gross_profit: t('gross_profit'), all_overhead: t('all_overhead'),
              result: t('result') };
  o.gm_pct = o.gross_profit / o.revenue;
  return o;
};
const TRIO = [roll('Support', [depts[0]]),
              roll('Production', [depts[1], depts[2]]),
              roll('Consulting', [depts[3], depts[4]])];

snipSlide({
  eyebrow: 'Financial performance', title: `The three departments, ${SHORT}`,
  snip: 'dept_three', source: FROM('Overhead Split', 'section 8, the three departments the P&L is built on'),
  commentary: (() => {
    const [sup, prod, cons] = TRIO;
    const word = (x) => (x.result >= 0 ? `made ${k(x.result)}k` : `lost ${k(-x.result)}k`);
    return [
      `This is how the P&L itself groups us. Production is Events and Video together. Consulting is Integration and Consulting together.`,
      `Production sold ${k(prod.revenue)}k against a ${k(prod.budget)}k plan, keeps ${pct(prod.gm_pct, 0)} of it, and ${word(prod)} after ${k(-prod.all_overhead)}k of running costs. It carries the month.`,
      `Support sold ${k(sup.revenue)}k against ${k(sup.budget)}k and ${word(sup)}. It keeps ${pct(sup.gm_pct, 0)}, which is well below where it normally runs.`,
      `Consulting sold ${k(cons.revenue)}k against a ${k(cons.budget)}k plan and keeps only ${pct(cons.gm_pct, 0)}, so it ${word(cons)}.`,
      `The three add to ${k(dTot.result)}k, the same as the five. Nothing is lost in the grouping.`,
    ];
  })(),
  notes: `The same month read three ways instead of five, which is how the Summary tab is actually built.\n\nProduction is the parent of Events and Video. Consulting is the parent of Integration and Consulting. Support has no children. Admin is not a trading department; its costs are spread across the three.\n\n` +
    TRIO.map((x) => `${x.name}: revenue ${money(x.revenue)} against ${money(x.budget)} budget. Gross profit ${money(x.gross_profit)} at ${pct(x.gm_pct)}. All overhead ${money(-x.all_overhead)}. Result ${money(x.result)}.`).join('\n') +
    `\n\nThe three total ${money(dTot.result)}, the operating result for the month, and the check row on the tab proves it against the five department version above.\n\nWhat the grouping shows that the five did not. Production, Events and Video together, sold ${money(TRIO[1].revenue)} at a ${pct(TRIO[1].gm_pct)} margin and carried ${money(-TRIO[1].all_overhead)} of running costs, leaving ${money(TRIO[1].result)}. That is the month. In the months it had volume it was always going to look like this, and ${MONTH} is the month it had volume.\n\nConsulting, Integration and Consulting together, sold ${money(Math.abs(TRIO[2].revenue - TRIO[2].budget))} ${TRIO[2].revenue >= TRIO[2].budget ? 'more' : 'less'} than plan and keeps only ${pct(TRIO[2].gm_pct)}, because most of it is resold equipment through Integration. Selling more of it will not change the result much.\n\nSupport is the line to ask about. It sold ${money(Math.abs(TRIO[0].revenue - TRIO[0].budget))} ${TRIO[0].revenue >= TRIO[0].budget ? 'above' : 'below'} plan at a ${pct(TRIO[0].gm_pct)} margin and came out at ${money(TRIO[0].result)}. It is normally the steady one, so the question for the room is whether that is timing on contract billing or something that carries into next month.`,
});

/* ---------------------------------- one P&L per department ---------------- */
// The pack's Dept P&L tab, pointed at each department in turn. Same columns
// every time: the month, the same month last year, and budget, each with a
// variance, so the three comparisons a department head asks for are all there.
const DEPT_STORY = {
  Support: (d) => ([
    `Billed ${k(d.revenue)}k against ${k(d.budget)}k, ${d.revenue >= d.budget ? 'ahead by' : 'short by'} ${k(Math.abs(d.revenue - d.budget))}k. Last ${LY_NAME} it billed ${k(d.ly)}k.`,
    `Kept ${pct(d.gm_pct, 0)} of what it sold. That is the lowest margin it has shown this year.`,
    `After its ${k(-d.all_overhead)}k share of running costs the month comes out at ${k(d.result)}k.`,
    `Of that share, ${money(staff[0].office)} is office wages. Support picks up the most because it works the most hours.`,
    `Year to date ${k(d.ytd_revenue)}k against a ${k(d.ytd_budget)}k plan, behind by ${k(d.ytd_budget - d.ytd_revenue)}k. ${k(d.budget - d.revenue)}k of that is this month.`,
  ]),
  Events: (d) => ([
    `Billed ${k(d.revenue)}k against ${k(d.budget)}k, ahead by ${k(d.revenue - d.budget)}k. Its largest month of the year.`,
    `Kept ${pct(d.gm_pct, 0)} of what it sold, which is ${k(d.gross_profit)}k of gross profit.`,
    `After ${k(-d.all_overhead)}k of running costs it made ${k(d.result)}k. It has carried the company this month.`,
    `${k(-d.alloc_overhead)}k of that charge is office cost, split on who uses the office rather than on what each department bills.`,
    `Year to date ${k(d.ytd_revenue)}k against a ${k(d.ytd_budget)}k plan. One month does not settle the year.`,
  ]),
  Video: (d) => ([
    `Billed ${k(d.revenue)}k against ${k(d.budget)}k, ${d.revenue >= d.budget ? 'ahead by' : 'short by'} ${k(Math.abs(d.revenue - d.budget))}k.`,
    `Keeps ${pct(d.gm_pct, 0)} of what it sells, the best margin in the business by a wide margin.`,
    `On ${k(d.revenue)}k of sales that is ${k(d.gross_profit)}k of gross profit, against ${k(-d.all_overhead)}k of running costs.`,
    `So it made ${k(d.result)}k. At this margin, volume in Video is worth more than volume anywhere else.`,
    `Year to date ${k(d.ytd_revenue)}k against ${k(d.ytd_budget)}k. The question is work coming in, not what we charge.`,
  ]),
  Integration: (d) => ([
    `Billed ${k(d.revenue)}k against a ${k(d.budget)}k plan, ahead by ${k(d.revenue - d.budget)}k.`,
    `Keeps only ${pct(d.gm_pct, 0)} of it. We buy the equipment and resell it, so most of the money goes to suppliers.`,
    `${k(d.revenue)}k of sales turned into ${k(d.gross_profit)}k. Do not read the revenue line as a good month.`,
    `It shows ${k(d.result)}k because it carries almost no running costs, ${k(-d.all_overhead)}k.`,
    `Year to date ${k(d.ytd_revenue)}k against ${k(d.ytd_budget)}k, ${d.ytd_revenue >= d.ytd_budget ? 'ahead by ' + k(d.ytd_revenue - d.ytd_budget) + 'k' : 'behind by ' + k(d.ytd_budget - d.ytd_revenue) + 'k'}. A good month does not fix that.`,
  ]),
  Consulting: (d) => ([
    `Billed ${k(d.revenue)}k against a ${k(d.budget)}k plan. There is almost no fee income here.`,
    `Carries ${k(-d.all_overhead)}k of running costs against ${k(d.revenue)}k of sales, so it loses ${k(-d.result)}k.`,
    `${k(-d.own_overhead)}k of that is its own spending. The rest is its share of the office.`,
    `Year to date it has billed ${k(d.ytd_revenue)}k against a ${k(d.ytd_budget)}k plan.`,
    `This is a decision for the board, not a forecast. Rebuild it, merge it into Integration, or close it.`,
  ]),
};


for (const name of ['Support', 'Events', 'Video', 'Integration', 'Consulting']) {
  const d = depts.find((x) => x.name === name);
  snipSlide({
    eyebrow: 'Financial performance', title: `${name}, ${SHORT}`,
    snip: `dept_${name}`,
    source: FROM('Dept P&L', `the department selector set to ${name}`),
    commentary: DEPT_STORY[name](d),
    notes: `${name} for ${MONTH}, read off the Dept P&L tab.\n\n` +
      `Sales ${money(d.revenue)} against a budget of ${money(d.budget)}, ${d.revenue >= d.budget ? 'ahead' : 'behind'} by ${money(Math.abs(d.revenue - d.budget))}. Last ${LY_NAME} it billed ${money(d.ly)}.\n\n` +
      `After direct costs of ${money(-d.cos)} it kept ${money(d.gross_profit)}, which is ${pct(d.gm_pct)} of what it sold.\n\n` +
      `Running costs charged to it were ${money(-d.all_overhead)}. ${money(-d.own_overhead)} of that is its own spending and ${money(-d.alloc_overhead)} is its share of the office, charged on the keys set in the pack.\n\n` +
      `That leaves ${money(d.result)} for the month.\n\n` +
      `Year to date it has billed ${money(d.ytd_revenue)} against a budget of ${money(d.ytd_budget)}.\n\n` +
      `On people: employment expense of ${money(staff.find((x) => x.name === name).total)}, of which ${money(staff.find((x) => x.name === name).departmental)} is its own staff and ${money(staff.find((x) => x.name === name).office)} is its share of office wages.` +
      (HOURS
        ? ` It put ${money(util.find((x) => x.name === name).chargeable)} hours onto jobs and ${money(util.find((x) => x.name === name).non_chargeable)} onto work we cannot bill, with ${money(util.find((x) => x.name === name).leave)} hours of leave.`
        : ` Hours for ${MONTH} have not been pasted on to the Utilisation tab yet, so there is no chargeable split for the month.`),
  });
}

snipSlide({
  eyebrow: 'Financial performance', title: 'Revenue against budget and last year',
  snip: 'rev_by_dept', source: FROM('Cost Analysis', 'revenue by department, the month'),
  commentary: (() => {
    const by = [...depts].sort((a, b) => (b.revenue - b.budget) - (a.revenue - a.budget));
    const v = (x) => `${k(x.revenue)}k against ${k(x.budget)}k`;
    return [
      `${by[0].name} billed ${v(by[0])}, ahead by ${k(by[0].revenue - by[0].budget)}k. That is where the month was made.`,
      `Integration billed ${v(depts[3])}, ahead by ${k(depts[3].revenue - depts[3].budget)}k, but at a ${pct(depts[3].gm_pct, 0)} margin. Revenue, not profit.`,
      `Support billed ${v(depts[0])}, short by ${k(depts[0].budget - depts[0].revenue)}k. That is unusual for contracted work and is worth a question.`,
      `Video billed ${v(depts[2])} and Consulting ${v(depts[4])}. Consulting has been short every month this year.`,
      `Company sales are ${pct(rev.month / rev.month_ly - 1, 0)} ${rev.month >= rev.month_ly ? 'up on' : 'down on'} last ${LY_NAME}, and we kept ${pct(gp.month / rev.month, 0)} of them against ${pct(gp.month_ly / rev.month_ly, 0)} then.`,
    ];
  })(),
  notes: `Revenue by department for ${MONTH}, actual against budget and against the same month last year.\n\n` +
    byRev.map((d) => `${d.name}: ${money(d.revenue)} against a budget of ${money(d.budget)} and ${money(d.ly)} last year.`).join('\n') +
    `\n\nCompany revenue ${money(rev.month)} against ${money(rev.month_budget)} budget and ${money(rev.month_ly)} last ${LY_NAME}.\n\nThere are two beats and they are not the same kind. Events is ahead of plan by ${money(depts[1].revenue - depts[1].budget)} at a ${pct(depts[1].gm_pct)} margin, so it turns into profit. Integration is ahead by ${money(depts[3].revenue - depts[3].budget)} at ${pct(depts[3].gm_pct)}, so it does not.\n\nThe line to watch is Support, ${money(depts[0].budget - depts[0].revenue)} under plan at a ${pct(depts[0].gm_pct)} margin. On contracted monthly work that is not what I would expect, and it is the reason the department result is negative.`,
});

/* ===================================================== 02 cost base */
section('02', 'What we spend', 'Where the money goes and which lines move',
  `Running costs are steady at about ${k(ohRun)} thousand a month. Spending is not the problem, but that is the number sales have to beat.`);

const cd = (t) => D.cost_dist.find((x) => x.label.includes(t));
snipSlide({
  eyebrow: 'What we spend', title: 'Operating cost distribution, year to date',
  snip: 'cost_dist', source: FROM('Cost Analysis', 'operating cost distribution'),
  commentary: [
    `Wages for people on jobs ${k(-cd('labour').amount)}k, and wages for everyone else ${k(-cd('Employment').amount)}k. Together that is most of what we spend.`,
    'Wages are what this business spends its money on. Everything else is small next to them.',
    `The ${k(-cd('Employment').amount)}k of wages for people not on jobs is the same every month. When sales drop this does not, so a bad month lands straight on the profit.`,
    `Equipment and services at ${k(-cd('equipment').amount)}k is the only cost that rises and falls with the work.`,
    'If we have to cut cost it is people or the office. Nothing else is big enough to matter.',
  ],
  notes: `Operating cost distribution for the year to date, cost of sales and overhead on one scale.\n\n` +
    D.cost_dist.filter((x) => Math.abs(x.amount) > 500).map((x) => `${x.label} ${money(Math.abs(x.amount))} at ${pct(x.share)}`).join('. ') +
    `.\n\nThe message is how concentrated it is. Wages for people on jobs and wages for everyone else are the bulk of it, and the second group is fixed for the year, which is why a quiet month drops straight off the profit.\n\nWhen the board asks what we can cut, the honest answer is staff numbers or the office. Trimming subscriptions and travel will not change this.`,
});

snipSlide({
  eyebrow: 'What we spend', title: 'Other operating expenses, year to date',
  snip: 'other_opex', source: FROM('Cost Analysis', 'other operating expenses'),
  commentary: (() => {
    const o = Object.fromEntries(D.other_opex.map((x) => [x.label, x]));
    const prem = Math.abs(o['Rent'].amount) + Math.abs(o['Rental Outgoings'].amount);
    return [
      `The office is the big one: rent ${k(-o['Rent'].amount)}k plus outgoings ${k(-o['Rental Outgoings'].amount)}k, ${k(prem)}k together.`,
      `IT support at ${k(-o['IT Network Service & Support'].amount)}k is next, and the only one that moves much month to month.`,
      `Everything below the top three adds to ${k(D.other_opex.slice(3).reduce((a, x) => a + x.amount, 0))}k across ${DONE} months.`,
      'There is nothing worth saving further down the list. Do not spend management time there.',
      'The office is the only real saving here, and that is a lease decision, not a budget one.',
    ];
  })(),
  notes: `Other operating expenses, year to date.\n\n` +
    D.other_opex.filter((x) => Math.abs(x.amount) > 100).map((x) => `${x.label} ${money(Math.abs(x.amount))} at ${pct(x.share)}`).join('. ') +
    `.\n\nRent and outgoings are fixed by the lease. IT support is the next largest and the only one that moves.\n\nThe practical conclusion is that savings do not come from here. Cutting this meaningfully means talking about the office lease.`,
});

snipSlide({
  eyebrow: 'What we spend', title: 'Labour mix, year to date',
  snip: 'labour_mix', source: FROM('Cost Analysis', 'labour mix, direct against indirect'),
  commentary: [
    'Labour on jobs splits two ways: our own staff, and contractors we bring in.',
    'Salaries, wages and super are most of it, and we are committed to paying them.',
    'Contractors are the only labour cost that goes up and down with the work.',
    'It rises when our own team runs out of hours, so it is the first place a capacity problem shows up.',
    'We pay for every hour worked. We only bill the hours spent on jobs. The difference is profit we never make.',
  ],
  notes: `Labour mix for the year to date, direct labour charged to cost of sales split by type.\n\n` +
    D.labour_mix.filter((x) => Math.abs(x.amount) > 500 && !/^all |total|check|in total/i.test(x.label))
      .map((x) => `${x.label} ${money(Math.abs(x.amount))}`).join('. ') +
    `.\n\nThe great majority is permanent payroll. Contractors are the only labour cost that rises and falls with the work.\n\nThis is the link to the utilisation section. We pay for hours worked and we bill for hours that are chargeable. Where those two diverge, margin goes with it.`,
});

snipSlide({
  eyebrow: 'What we spend', title: 'Employment expense, office and departmental',
  snip: 'staff_split', source: FROM('Overheads MoM', 'employment expense split, the month'),
  commentary: (() => {
    const tot = staff.reduce((a, x) => a + x.office, 0);
    return [
      `${money(tot)} of office wages this month, and all of it is charged out to the departments.`,
      `Support picks up ${pct(staff[0].office / tot, 0)} of it, based on its share of hours worked over the last twelve months.`,
      `Events pays more for its own staff, ${money(staff[1].departmental)}, than it does for office staff.`,
      'Integration picks up almost no office wages, because almost no hours were booked to it.',
      'This is the line that used to be typed by hand. It now calculates from the labour hours report.',
    ];
  })(),
  notes: `Employment expense split between office staff and departmental staff for ${MONTH}.\n\nDepartmental staff is the employment expense on each department's own cost centre: ` +
    staff.map((x) => `${x.name} ${money(x.departmental)}`).join(', ') +
    `.\n\nOffice staff is the admin employment pot of ${money(staff.reduce((a, x) => a + x.office, 0))}, charged across the departments on their share of full time equivalent: ` +
    staff.map((x) => `${x.name} ${money(x.office)}`).join(', ') +
    `.\n\nFTE is rolling twelve month worked hours on that cost centre divided by ${money(D.annual_hours)} hours, being 253 working days at eight. That is the method in the training recording.\n\nThe change worth noting is that this no longer relies on a typed figure. The pack calculates it from the Labour Hours Report, so it reproduces every month without anyone deciding it.`,
});

/* ===================================================== 03 customers and cash */
section('03', 'Customers and cash', 'Who pays us, how much, and how slowly',
  `Revenue is concentrated and cash is slow. Days sales outstanding is ${D.dso.days.toFixed(1)} against a ${D.dso.target} day target, which is ${money(D.dso.locked)} of cash sitting in the debtor book above where it should be.`);

snipSlide({
  eyebrow: 'Customers and cash', title: 'Top ten customers, year to date',
  snip: 'top10', source: FROM('Top 10 Clients', 'top ten customers by revenue, the top three in red'),
  commentary: (() => {
    const t = D.top10, tt = D.top10_total;
    const top3 = t.slice(0, 3).reduce((a, x) => a + x.ytd, 0);
    return [
      `The three red bars are ${pct(top3 / rev.ytd, 0)} of everything we sold this year.`,
      `All ten together are ${pct(tt.top_ten_share, 0)}. Every other customer we have is the remaining ${pct(tt.rest_share, 0)}.`,
      `${t[0].name} at ${money(t[0].ytd)} is ${pct(t[0].share, 0)} on its own. Losing them would wipe out more than the whole year's profit plan.`,
      `${t[0].name} billed ${money(t[0].month)} in ${MONTH} alone, which is ${pct(t[0].month / rev.month, 0)} of everything we sold in the month.`,
      `After the top five it falls away fast. Number ten is only ${money(t[9].ytd)}.`,
    ];
  })(),
  notes: `Top ten customers by revenue, year to date, with the top three called out in red.\n\n` +
    D.top10.map((x) => `${x.rank}. ${x.name}: ${money(x.ytd)} for the year, ${pct(x.share)} of the total, ${money(x.month)} of it in ${MONTH}.`).join('\n') +
    `\n\nThe ten together are ${money(D.top10_total.top_ten_ytd)}, or ${pct(D.top10_total.top_ten_share)} of everything we sold. Every other customer on the books makes up the remaining ${money(D.top10_total.rest_ytd)}.\n\nThe three red bars are the point of the slide. They are ${pct(D.top10.slice(0, 3).reduce((a, x) => a + x.ytd, 0) / rev.ytd)} of the year between them.\n\nThe concentration moved this month. ${D.top10[0].name} billed ${money(D.top10[0].month)} in ${MONTH}, which is ${pct(D.top10[0].month / rev.month, 0)} of the month on its own, and is now ${pct(D.top10[0].share)} of the year.\n\nThe difference that matters is contract work against one-off projects. Contracted support renews every month. A project bills once and stops. The pack does not separate the two, so somebody has to mark which of the top three is which before this goes to the board. Counting them as the same kind of income is how a forecast goes wrong.\n\nWhat I am recommending: a named owner for each of the top three before the half year, and the new business target set on customers four to twenty rather than on the ones we already have.`,
});

snipSlide({
  eyebrow: 'Customers and cash', title: 'How long customers take to pay',
  snip: 'dso_calc', source: FROM('DSO', 'the calculation, built off the Xero age analysis'),
  commentary: (() => {
    const d = D.dso;
    return [
      `Customers take ${d.days.toFixed(0)} days to pay us. We want ${d.target}.`,
      `That is ${d.above.toFixed(0)} days of sales sitting in the debtor book instead of the bank.`,
      `In money, ${money(d.locked)} that should be ours already.`,
      `Every single day we pull that back brings in ${money(d.receivables / d.days)}.`,
      `${D.dso.days > 45 ? 'And it is getting slower, not faster.' : 'The direction is right, keep it going.'} This is the fastest cash available to the business. Nothing else comes close.`,
    ];
  })(),
  notes: `How long customers take to pay, at ${MONTH_END}.\n\nWe are owed ${money(D.dso.receivables)}. Across the last three months we sold ${money(D.dso.revenue_3m)}, which over ${D.dso.days_in || 92} days is ${money(D.dso.receivables / D.dso.days)} of sales a day. Divide what we are owed by that and it takes customers ${D.dso.days.toFixed(1)} days to pay us.\n\nThe target is ${D.dso.target} days. We are ${D.dso.above.toFixed(1)} days over it. If we collected on time we would be owed ${money(D.dso.at_target)} instead of ${money(D.dso.receivables)}, so ${money(D.dso.locked)} of cash that should already be ours is sitting in the debtor book.\n\nPut that next to the profit position. The ${DONE} months to date made ${money(eb.ytd)} of EBITDA. There is ${money(D.dso.locked)} of our own money uncollected. Collecting it is worth more than anything we can trade our way to between now and Christmas, and it costs nothing but phone calls.\n\nEvery day we take off the number brings in ${money(D.dso.receivables / D.dso.days)}.`,
});

snipSlide({
  eyebrow: 'Customers and cash', title: 'Who owes us the most',
  snip: 'dso_who', source: FROM('DSO', 'the ten largest amounts owed at month end'),
  commentary: (() => {
    const d = D.dso, deb = D.debtors;
    return [
      `${deb[0].name} owes ${money(deb[0].amount)} on its own, ${pct(deb[0].share, 0)} of everything owed to us.`,
      `With ${deb[1].name} at ${money(deb[1].amount)}, two customers hold ${pct((deb[0].share + deb[1].share), 0)} of the book.`,
      `${pct(d.overdue_share, 0)} of what we are owed is already late, ${money(d.overdue)}.`,
      `Chase those two first. Everything below them is small by comparison.`,
      `A general collections push is not the answer. Two phone calls are.`,
    ];
  })(),
  notes: `The ten largest amounts owed at ${MONTH_END}.\n\n` +
    D.debtors.slice(0, 5).map((x) => `${x.name}: ${money(x.amount)}, ${pct(x.share)} of the total.`).join('\n') +
    `\n\nThe ageing: ` + D.ageing.filter((x) => x.amount > 0).map((x) => `${x.bucket} ${money(x.amount)} at ${pct(x.share)}`).join(', ') +
    `. Late in total ${money(D.dso.overdue)}, or ${pct(D.dso.overdue_share)} of the book.\n\nThe action here is narrow and that is the point. Two customers hold most of the balance. A targeted chase on those two is worth more than a general collections exercise across everyone, and it is the fastest money available to us.`,
});

/* ===================================================== 04 people and utilisation */
// Both inputs are monthly pastes: the hours on the Utilisation tab, and her own
// Utilisation Report figures in deck/data/util_her_basis.json. Until the
// reporting month is in both there is nothing to show, so the section is left
// out rather than printed at nil.
if (SHOW_PEOPLE) {
section('04', 'People and utilisation', 'Chargeable hours and the profit they produce',
  `The company billed ${money(HU.Grand.ch)} hours in ${MONTH} out of ${money(HU.Grand.workable)} it was available for, which is ${pct(HU.Grand.pct, 0)}. Where chargeable hours fall short, margin follows.`);

// Her own Chargeable % comes first, because it is the number the business
// already uses. Billable hours over the hours the person was available for,
// leave and public holidays included in the bottom of the fraction.
// One slide instead of four. The three month history and the hours and FTE
// detail are on the Utilisation tab; what the board needs in the room is the
// month, per group, against target.
snipSlide({
  eyebrow: 'People and utilisation', title: `Chargeable per cent, ${SHORT}`,
  snip: 'util_month', source: FROM('Utilisation', 'chargeable per cent against target, the Utilisation Report basis'),
  commentary: [
    `The monthly report's own calculation: billable hours divided by the hours each person was available for, so leave counts against it.`,
    `Support ${pct(HU.Support.pct, 0)} against an ${pct(HU.Support.target, 0)} target, ${pct(HU.Support.over_under, 1)} over. The only group ahead.`,
    `Production ${pct(HU.Production.pct, 0)} against ${pct(HU.Production.target, 0)}. Leave took ${pct(HU.Production.leave_pct, 0)} of its available hours.`,
    `The company is ${pct(HU.Grand.pct, 0)} against ${pct(HU.Grand.target, 0)}, so ${pct(-HU.Grand.over_under, 1)} under.`,
  ],
  notes: `Chargeable per cent for ${MONTH}, calculated exactly as the monthly Utilisation Report calculates it.\n\nThe formula is billable hours divided by Workable Business Hours, the hours each person was available for in the month. Leave and public holidays stay in the bottom of the fraction, so a team that takes its leave shows a lower percentage. That is the business's own definition and this pack now uses it.\n\nSupport ${pct(HU.Support.pct)} against an ${pct(HU.Support.target)} target, ${pct(HU.Support.over_under)} over. Production ${pct(HU.Production.pct)} against ${pct(HU.Production.target)}, ${pct(-HU.Production.over_under)} under. Consulting ${pct(HU.Consulting.pct)} against ${pct(HU.Consulting.target)}, ${pct(-HU.Consulting.over_under)} under. The company ${pct(HU.Grand.pct)} against ${pct(HU.Grand.target)}, ${pct(-HU.Grand.over_under)} under.\n\nSupport has cleared its own target in all three closed months, by ${pct(HER['Jul-26'].Support.over_under)}, ${pct(HER['Aug-26'].Support.over_under)} and ${pct(HU.Support.over_under)}.\n\nLeave is the story on Production: ${money(HU.Production.lv)} hours out of ${money(HU.Production.workable)} available, which is ${pct(HU.Production.leave_pct)} of the month. Across the company leave took ${pct(HU.Grand.leave_pct)} of available hours against ${pct(HER['Aug-26'].Grand.leave_pct)} in August.\n\nThree person groups, not five. The report groups by the department a person belongs to and nobody is assigned to Video or Integration; their work is carried out by Production and Consulting people.\n\nWhat is behind this slide. Workable Business Hours and the target are pastes, because neither can be calculated in the pack: the first needs each group's headcount and who is part time, the second is set per person and averaged up. The over or under reading and the leave share are calculated and reproduce the report's own columns exactly. The three month history sits on the Utilisation tab at rows 120 to 164, with two check rows against the report's Grand Total.\n\nA correction to earlier versions of this pack. It divided billable hours by hours actually worked, leaving leave out of the sum, which put Support at 99.5 per cent and made it look like a coding error. It was not.\n\nOT and TIL is on the report but is not in the pack. Its August figure for Support reads 365.62, which cannot be hours in a month, so that column needs checking on the source sheet before anybody quotes it.`,
});

snipSlide({
  eyebrow: 'People and utilisation', title: `How each department spent its hours, ${SHORT}`,
  snip: 'util_chart', source: FROM('Utilisation', 'hours by department, each bar scaled to its own total'),
  commentary: (() => {
    const sup = util[0], ev = util[1], con = util[4];
    return [
      `Each bar is scaled to that department's own hours, so you can compare the mix. Support works about a hundred times the hours Integration does.`,
      `Support is almost entirely billable, which is why it is the only department that makes money.`,
      `Events put ${money(ev.leave)} hours into leave and only ${money(ev.chargeable)} onto jobs. The people are there, the work is not.`,
      `Consulting put ${money(con.non_chargeable)} hours into work we cannot bill against ${money(con.chargeable)} we can.`,
      `We pay for every hour. We only bill the ones on jobs. The grey and light teal are the gap.`,
    ];
  })(),
  notes: `How each department spent its hours in ${MONTH}.\n\nRead the bars as a mix, not as a size. Each one is scaled to that department's own total, because Support works around a hundred times the hours Integration does and on a straight scale four of the five departments disappear. The number on each segment is the real hours.\n\n` +
    util.map((u) => `${u.name}: ${money(u.chargeable)} hours on jobs, ${money(u.non_chargeable)} not on jobs, ${money(u.leave)} leave. ${pct(u.util_pct)} of worked hours were billable.`).join('\n') +
    `\n\nAcross the five departments, ${money(utilTot.c)} hours on jobs out of ${money(utilTot.c + utilTot.n)} worked.\n\nOn the Support bar: the small slice of non-billable time comes from the regrouping, not from bad coding. The monthly report shows Support with no non-billable hours in ${MONTH}. This view regroups by the job the work was for, so a little of other people's non-billable time lands here.\n\nEvents is the real operational issue: ${money(util[1].leave)} hours of leave on a small team takes out most of the month's capacity.\n\nAnd the prize. Support earns about ${money(depts[0].revenue / util[0].chargeable - Math.abs(depts[0].cos) / util[0].chargeable)} of margin on every hour it bills. Five more points of billable time across Support and Events, at the rate each already earns, is worth roughly ${money((depts[0].revenue + depts[1].revenue) * 0.05)} a month and costs nothing extra. It is the largest controllable item in this pack.`,
});

snipSlide({
  eyebrow: 'People and utilisation', title: `Hours, the three departments, ${SHORT}`,
  snip: 'util_three', source: FROM('Utilisation', 'hours rolled up to the three departments, each bar scaled to its own total'),
  commentary: (() => {
    const prod = { c: util[1].chargeable + util[2].chargeable,
                   n: util[1].non_chargeable + util[2].non_chargeable,
                   l: util[1].leave + util[2].leave };
    const cons = { c: util[3].chargeable + util[4].chargeable,
                   n: util[3].non_chargeable + util[4].non_chargeable,
                   l: util[3].leave + util[4].leave };
    return [
      `The same picture grouped the way the P&L is. Production is Events and Video. Consulting is Integration and Consulting.`,
      `Production put ${money(prod.c)} hours onto jobs and ${money(prod.n)} onto work we cannot bill, with ${money(prod.l)} hours of leave on top.`,
      `Consulting put ${money(cons.c)} onto jobs against ${money(cons.n)} we cannot bill. Two thirds of its paid hours earn nothing.`,
      `Support is the only bar that is almost all dark. That is why it is the only department that makes money.`,
      `Each bar is scaled to its own hours. Support works about four times the hours of the other two together.`,
    ];
  })(),
  notes: `The hours mix for ${MONTH}, rolled up to the three departments the P&L is built on.\n\nProduction is Events and Video together. Consulting is Integration and Consulting together. Support stands alone.\n\nRead each bar as a mix, not a size. Each is scaled to its own total, because Support works far more hours than the other two combined.\n\nThe pattern is the whole argument of this pack. Support is almost entirely billable and it is the only department that covers its costs. Production and Consulting both spend roughly half their paid hours, or more, on work no customer pays for, and both lose money.\n\nSupport reads ${pct(HU.Support.pct)} on the monthly report's own basis, which counts leave against it. That is the figure to quote in the room.`,
});

} else {
  console.warn('  people and utilisation section skipped: no hours pasted for ' + MONTH);
}

/* ===================================================== 05 budget and outlook */
section(SHOW_PEOPLE ? '05' : '04', 'Budget and outlook', 'What the rest of the year looks like and what we change',
  `${DONE} months are closed and the remaining ${LEFT} are still at original budget. That is the decision in front of the board.`);

snipSlide({
  eyebrow: 'Budget and outlook', title: 'Revenue by month against budget',
  snip: 'income_month', source: FROM('Cost Analysis', 'total income by month, actual to date then budget'),
  commentary: (() => {
    const m = D.income_by_month;
    const a = m.slice(0, DONE).reduce((x, y) => x + y.actual, 0);
    const b = m.slice(0, DONE).reduce((x, y) => x + y.budget, 0);
    const best = Math.max(...m.slice(0, DONE).map((x) => x.actual));
    return [
      `${DONE} months in, we are ${k(Math.abs(a - b))}k ${a >= b ? 'ahead of' : 'behind'} budget on revenue, at ${pct(a / b, 0)} of plan.`,
      `${SHORT} closed that gap. It billed ${k(m[DONE - 1].actual)}k against ${k(m[DONE - 1].budget)}k and is our best month of the year.`,
      `The remaining ${LEFT} months are untouched original budget. They are not a forecast.`,
      'December carries a planned loss. The business closes for two weeks, so that is correct and should stay.',
      'Recommend a rolling forecast from here: rebuild Consulting on real capacity and reset the margin assumptions.',
    ];
  })(),
  notes: (() => {
    const m = D.income_by_month;
    const a = m.slice(0, DONE).reduce((x, y) => x + y.actual, 0);
    const b = m.slice(0, DONE).reduce((x, y) => x + y.budget, 0);
    return `Revenue by month, actual for the ${DONE} closed months and original budget for the rest.\n\nYear to date actual ${money(a)} against a budget of ${money(b)}, ${a >= b ? 'ahead by' : 'behind by'} ${money(Math.abs(b - a))}, which is ${pct(a / b, 1)} of plan.\n\n${MONTH} billed ${money(m[DONE - 1].actual)} against a budget of ${money(m[DONE - 1].budget)}. It is the largest month of the year and it took the year to date from well behind plan to close to it.\n\nThe ${LEFT} months from ${m[DONE] ? new Date(m[DONE].month).toLocaleString('en-AU', { month: 'long' }) : 'here'} are still at the budget set before the year started.\n\nTwo things for the board. The variance to date is concentrated in Consulting and Video, which were budgeted on a resourcing level that does not exist. And the margin assumption in the budget is above what the business has delivered for two years, so even a month that hits the revenue number misses the profit number. ${MONTH} is exactly that: revenue ${money(m[DONE - 1].actual - m[DONE - 1].budget)} over plan and gross profit ${money(Math.abs(gp.month - gp.month_budget))} under it.\n\nMy recommendation is a rolling forecast from here, rebuilding Consulting around actual capacity and resetting gross margin to delivered rates. December should stay as budgeted: the business closes for two weeks and the loss is real.`;
  })(),
});

/* ===================================================== summary and actions */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Summary and actions', 'What we agreed');
  // Dates run off the reporting month, so they move with the pack instead of
  // being typed. Day 0 of the next month is the last day of this one.
  const base = new Date(D.income_by_month[DONE - 1].month);
  const due = (addMonths, day) => {
    const d = day ? new Date(base.getFullYear(), base.getMonth() + addMonths, day)
                  : new Date(base.getFullYear(), base.getMonth() + addMonths + 1, 0);
    return d.toLocaleDateString('en-AU', { day: '2-digit', month: 'short', year: '2-digit' })
            .replace(/\./g, '');
  };
  const top3 = D.top10.slice(0, 3).reduce((a, x) => a + x.ytd, 0);
  const acts = [
    ['P&L', `Reforecast FY27 on margin, not revenue. Revenue is ${pct(rev.ytd / rev.ytd_budget, 0)} of plan and gross profit only ${pct(gp.ytd / gp.ytd_budget, 0)}. Break even needs ${k(breakEven)}k a month at the ${pct(ytdMargin)} margin we deliver and we are running at ${k(revRun)}k.`, 'GC and FC', due(1)],
    ['Overhead keys', `Office wages of ${money(staff.reduce((a, x) => a + x.office, 0))} are charged on share of hours worked, so Support and Events carry nearly all of it whatever they bill. Test the keys for FY28, do not change them mid year.`, 'GC and FC', due(2)],
    ['Consulting', `${k(-depts[4].all_overhead)}k of running costs against ${k(depts[4].revenue)}k of fees, and ${k(depts[4].ytd_revenue)}k billed year to date against a ${k(depts[4].ytd_budget)}k plan. Decide whether the department is rebuilt, merged into Integration, or closed.`, 'GC', due(2)],
    ['Customers', `${D.top10[0].name} alone is ${pct(D.top10[0].share)} of the year and billed ${money(D.top10[0].month)} in ${MONTH}. Top three are ${pct(top3 / rev.ytd)}. Agree a named retention plan and mark which of them is contract and which is project.`, 'GC and Sales', due(2, 14)],
    ['Cash', `DSO ${D.dso.days.toFixed(1)} days against a ${D.dso.target} day target, up from the end of the quarter. Getting to target releases ${k(D.dso.locked)}k. Chase ${D.debtors[0].name} and ${D.debtors[1].name} first.`, 'FC', due(1)],
    ['Support margin', `Support billed ${k(depts[0].revenue)}k against ${k(depts[0].budget)}k at a ${pct(depts[0].gm_pct)} margin and came out at ${k(depts[0].result)}k. Confirm whether that is billing timing or a real change in the contract mix.`, 'FC and Ops', due(1)],
  ];
  if (HU) {
    const sup = util[0], ev = util[1];
    const gain = (depts[0].revenue + depts[1].revenue) * 0.05;
    acts.push(['Utilisation', `Chargeable per cent is ${pct(HU.Grand.pct, 0)} against the report's own ${pct(HU.Grand.target, 0)} target. Five more points on Support and Production is worth about ${k(gain)}k a month at no extra cost.`, 'Ops and FC', due(2)]);
  } else {
    acts.push(['Utilisation', `${MONTH} hours are not in the pack yet. Paste the Utilisation Report and the month's hours so the chargeable per cent can be reported with the rest of the month.`, 'Ops and FC', due(1)]);
  }
  s.addTable([headRow(['Area', 'What we agreed', 'Owner', 'Due'], 'left'),
    ...acts.map(([a, w, o, dd]) => [
      { text: a, options: { fontSize: 8, bold: true, color: C.navy, align: 'left', valign: 'top', margin: [3, 5, 3, 5] } },
      { text: w, options: { fontSize: 8, color: C.body, align: 'left', valign: 'top', margin: [3, 5, 3, 5] } },
      { text: o, options: { fontSize: 8, color: C.body, align: 'left', valign: 'top', margin: [3, 5, 3, 5] } },
      { text: dd, options: { fontSize: 8, color: C.body, align: 'left', valign: 'top', margin: [3, 5, 3, 5] } },
    ])], {
    x: 0.55, y: 1.58, w: 8.90, colW: [1.00, 5.75, 1.05, 1.10],
    border: { type: 'solid', color: C.rule, pt: 0.5 }, autoPage: false,
  });
  s.addNotes(`${acts.length} actions, each with an owner and a date.\n\nThe first one matters most: reforecasting on margin rather than revenue. ${MONTH} is the proof of why. Revenue beat budget by ${money(rev.month - rev.month_budget)} and gross profit still came in ${money(Math.abs(gp.month - gp.month_budget))} under it.\n\nThe Consulting decision cannot keep being deferred. It has carried running costs against almost no fees every month this year.\n\nThe cash action is the only one that produces money now, and it has got worse: ${D.dso.days.toFixed(1)} days against a ${D.dso.target} day target, with ${money(D.dso.locked)} sitting in the debtor book above where it should be.\n\nThe dates are set off the reporting month. Check them against your own calendar before this goes out.`);
}

/* ===================================================== thank you */
{
  const s = pres.addSlide({ masterName: 'CTS_DARK' });
  [{ x: 5.5, y: 0.25, d: 2.5, c: C.tealLt }, { x: 6.8, y: 1.35, d: 2.1, c: C.blue }, { x: 5.8, y: 2.35, d: 2.6, c: C.red }]
    .forEach((o, i) => s.addShape('ellipse', { x: o.x, y: o.y, w: o.d, h: o.d, fill: { color: o.c, transparency: 18 }, line: { type: 'none' }, objectName: `closing circle ${i}` }));
  s.addText('cts', { x: 0.55, y: 0.45, w: 2, h: 0.7, isTextBox: true, margin: 0, fontSize: 34, bold: true, color: C.white });
  s.addText('Seamless AV', { x: 0.60, y: 0.98, w: 2, h: 0.25, isTextBox: true, margin: 0, fontSize: 8.5, color: C.grey });
  s.addText('Thank you', { x: 0.55, y: 2.30, w: 5, h: 0.8, isTextBox: true, margin: 0, fontSize: 38, bold: true, color: C.white, valign: 'middle' });
  s.addText('Questions and discussion', { x: 0.55, y: 3.10, w: 5, h: 0.4, isTextBox: true, margin: 0, fontSize: 13, color: C.grey, valign: 'middle' });
  s.addNotes(`Close on the two numbers that matter: break even needs ${money(breakEven)} a month at the ${pct(ytdMargin)} margin we actually deliver, and we are running at ${money(revRun)}. ${revRun >= breakEven ? 'We are above that line now, and ' + MONTH + ' is what put us there. It was volume, not margin, so the line holds only while the volume does.' : 'We are still below that line.'} Everything else in the pack is detail underneath that.`);
}

module.exports = { pres };
