const B = require('./build.js');
const L = require('./lib.js');
const { C, DEPT } = require('./theme.js');
const { head, commentary, headRow, row, money, k, pct, signed, varColour } = L;
const { placeSnip, source } = require('./snips.js');
const { EB } = require('./part0.js');
const HER = require('../data/util_her_basis.json');   // her own report's Chargeable %
const HU = HER['Sep-26'];
const { pres, D, P, rev, gp, oh, eb, np, depts, dTot, byRev, util, utilTot, staff,
        ytdMargin, ohRun, breakEven, revRun, section, MONTH, SHORT } = B;

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
  const box = wide ? { x: 0.55, y: 1.52, w: 8.90, h: 2.45 } : { ...BOX, ...(o.box || {}) };
  const r = placeSnip(s, o.snip, box);
  source(s, o.source, r.bottom + 0.08, box.w, box.x);
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
  `Three months of FY27 are closed. Revenue is running at ${k(revRun)} thousand a month and break even needs ${k(breakEven)} thousand at the margin we actually deliver. That is the whole problem in one line.`);

snipSlide({
  eyebrow: 'Financial performance', title: 'P&L summary, income',
  snip: 'pl_income', source: FROM('Summary', 'income by department, the three closed months and year to date'),
  commentary: [
    `We need ${k(breakEven)}k of sales a month just to cover our costs. We are averaging ${k(revRun)}k.`,
    `${SHORT} sold ${k(rev.month - rev.month_budget)}k more than planned and still lost ${k(-eb.month)}k. We sold enough. We did not keep enough of it.`,
    `We kept ${pct(gp.month / rev.month, 0)} of every dollar sold. The plan said ${pct(gp.month_budget / rev.month_budget, 0)}, but last September we kept ${pct(gp.month_ly / rev.month_ly, 0)}.`,
    `Running costs came in at ${k(-oh.month)}k against ${k(-oh.month_budget)}k planned. Spending is not the problem.`,
    `Profit for the quarter is ${k(eb.ytd)}k against ${k(eb.ytd_budget)}k planned. The plan needs redoing around what we keep, not what we sell.`,
  ],
  notes: `P&L summary, straight off the Summary tab with the outline shut, so this is the view on screen.\n\nRevenue ${money(rev.month)} against a budget of ${money(rev.month_budget)}, ahead by ${money(rev.month - rev.month_budget)}. Against last September at ${money(rev.month_ly)} we are up ${pct(rev.month / rev.month_ly - 1, 0)}.\n\nGross profit ${money(gp.month)} at ${pct(gp.month / rev.month)}, against a budget of ${pct(gp.month_budget / rev.month_budget)} and last September at ${pct(gp.month_ly / rev.month_ly)}. The budget margin for this month was unusually low, so beating it is not the achievement it looks like.\n\nOverheads ${money(-oh.month)} against ${money(-oh.month_budget)} budget. Spending is behaving.\n\nEBITDA ${money(eb.month)} against a budgeted ${money(eb.month_budget)}. Net profit ${money(np.month)} after depreciation and amortisation of ${money(-P.da.month)}.\n\nYear to date revenue ${money(rev.ytd)} against ${money(rev.ytd_budget)}, behind by ${money(rev.ytd_budget - rev.ytd)}. EBITDA ${money(eb.ytd)} against ${money(eb.ytd_budget)}.\n\nThe line to leave in the room: at ${pct(ytdMargin)} gross margin we need ${money(breakEven)} of revenue every month to cover overhead. We are averaging ${money(revRun)}. That gap is structural and selling more at this margin will not close it.`,
});

snipSlide({
  eyebrow: 'Financial performance', title: 'P&L summary, cost and profit',
  snip: 'pl_result', source: FROM('Summary', 'cost of sales, overheads and the profit lines'),
  commentary: [
    `Operating profit, or EBITDA, is ${k(EB.q1_actual)}k for the quarter against ${k(EB.q1_budget)}k planned. We are ${k(-EB.shortfall)}k short, which is ${pct(-EB.shortfall / EB.q1_budget, 0)} of it.`,
    `The plan expected ${pct(EB.q1_budget / EB.fy_budget, 0)} of the year's profit in these first three months. It did not happen, so the whole year now rides on the nine months left.`,
    `To still hit the ${k(EB.fy_budget)}k target those nine months have to make ${k(EB.still_needed)}k. The plan only ever asked them for ${k(EB.budget_assumed_rest)}k.`,
    `${k(EB.still_needed)}k is ${(EB.still_needed / EB.fy26_actual).toFixed(1)} times what the whole of last year made, which was ${k(EB.fy26_actual)}k. We will not get the year back.`,
    `The problem is what we keep, not what we sell. Sales hit ${pct(rev.ytd / rev.ytd_budget, 0)} of plan. What was left after direct costs hit only ${pct(gp.ytd / gp.ytd_budget, 0)}.`,
  ],
  notes: `The risk in one slide: EBITDA against what was budgeted.\n\nThe quarter produced ${money(EB.q1_actual)} of EBITDA against a budget of ${money(EB.q1_budget)}. That is a miss of ${money(-EB.shortfall)}, or ${pct(-EB.shortfall / EB.q1_budget, 1)}.\n\nWhy it matters more than the number suggests. The FY27 budget is ${money(EB.fy_budget)} of EBITDA for the year, and it put ${money(EB.q1_budget)} of that, ${pct(EB.q1_budget / EB.fy_budget, 0)}, into the first quarter. The year was planned to be earned early. It was not, so the remaining nine months now have to carry the whole thing.\n\nTo still land on ${money(EB.fy_budget)} the nine months from October must produce ${money(EB.still_needed)}, which is ${money(EB.per_month_needed)} a month. The budget itself only asked those nine months for ${money(EB.budget_assumed_rest)}, or ${money(EB.per_month_assumed)} a month. We would need ${(EB.still_needed / EB.budget_assumed_rest).toFixed(1)} times the monthly profit the plan assumed, in the months the plan already treated as the quiet ones.\n\nPut it against last year and it is clearer. The whole of FY26 delivered ${money(EB.fy26_actual)} of EBITDA. The next nine months would need ${money(EB.still_needed)}, which is ${(EB.still_needed / EB.fy26_actual).toFixed(1)} times a full year of last year's earnings. My view is that the full year EBITDA budget is not recoverable and the board should be told that now rather than at the half year.\n\nThe cause is margin, not sales. Revenue for the quarter came in at ${pct(rev.ytd / rev.ytd_budget, 1)} of budget, which is not a bad result on its own. Gross profit came in at ${pct(gp.ytd / gp.ytd_budget, 1)} of budget. The gap between those two percentages is the whole problem: we sold reasonably well and kept very little of it. September is the clearest example, where revenue beat budget by ${money(rev.month - rev.month_budget)} and the month still lost ${money(-eb.month)} of EBITDA.\n\nNothing in what we spend can soak that up. Overheads ran at ${money(-oh.ytd)} for the quarter, ${money(ohRun)} a month, and came in ${money(oh.ytd - oh.ytd_budget)} under budget. Spending is not the problem and there is nothing more to cut that would matter. Wages and the office are what we spend, and both are locked in for the year. So when we keep less of each sale it goes straight off the profit.\n\nWhere that leaves break even. At the ${pct(ytdMargin)} gross margin we actually deliver, the business needs ${money(breakEven)} of revenue a month to cover overhead. We are running at ${money(revRun)}. The business is trading at break even, not at a profit, and it is doing so in the quarter the budget expected to be the strongest.\n\nTwo things make the next two quarters harder, not easier. December is budgeted at a loss because the business closes for two weeks, so Q2 cannot make the gap up. And with EBITDA at break even there is no earnings buffer, while ${money(D.dso.locked)} of cash sits in debtors above the ${D.dso.target} day target.\n\nThe one piece of genuine good news: last year's first quarter was a loss of ${money(-EB.fy26_q1)}. This quarter is ${money(EB.q1_actual - EB.fy26_q1)} better than that. The business is improving year on year. It is the budget that was set above what the business can deliver, and that is the distinction to draw in the room.\n\nWhat I am recommending. Reforecast now, on margin rather than revenue. Rebuild Consulting around the capacity that actually exists. Reset the gross margin assumption to the rate delivered over two years rather than the rate in the plan. Hold December as budgeted. And give the board a revised full year EBITDA they can rely on, because the one in front of them cannot be met.`,
});

snipSlide({
  eyebrow: 'Financial performance', title: `Department result, ${SHORT}`,
  snip: 'dept_result', source: FROM('Overhead Split', 'section 7, department result'),
  commentary: [
    `Support is the only department earning enough to cover its share of running costs. It made ${k(depts[0].result)}k on ${k(depts[0].revenue)}k of sales.`,
    `Events is charged ${k(-depts[1].all_overhead)}k of running costs and only billed ${k(depts[1].revenue)}k. It keeps ${pct(depts[1].gm_pct, 0)} of what it sells and still loses ${k(-depts[1].result)}k.`,
    `Integration billed ${k(depts[3].revenue)}k against a ${k(depts[3].budget)}k plan but keeps only ${pct(depts[3].gm_pct, 0)} of it, so it made just ${k(depts[3].result)}k.`,
    `Consulting is charged ${k(-depts[4].all_overhead)}k of running costs and billed ${k(depts[4].revenue)}k. That needs a decision, not a forecast.`,
    `The five departments add up to the company result exactly. The check row on the tab proves it.`,
  ],
  notes: `Every department against budget with the full overhead charge, read off section 7 of the Overhead Split tab.\n\n` +
    depts.map((x) => `${x.name}: revenue ${money(x.revenue)} against ${money(x.budget)} budget. Gross profit ${money(x.gross_profit)} at ${pct(x.gm_pct)}. Overhead ${money(-x.all_overhead)}, of which ${money(-x.alloc_overhead)} is allocated admin. Result ${money(x.result)}.`).join('\n') +
    `\n\nThe five departments total ${money(dTot.result)}, which is net profit for the month.\n\nTwo things to flag. Integration sold ${money(depts[3].revenue - depts[3].budget)} more than planned and it turned into almost no profit, because we keep only ${pct(depts[3].gm_pct)} on reselling equipment. And Events is charged ${money(-depts[1].alloc_overhead)} of office costs on ${money(depts[1].revenue)} of sales. That charge is split on who uses the office, and Events is not the department sitting in it.`,
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
    return [
      `This is how the P&L itself groups us. Production is Events and Video together. Consulting is Integration and Consulting together.`,
      `Support made ${k(sup.result)}k on ${k(sup.revenue)}k of sales. It is the only one of the three that covers its own running costs.`,
      `Production sold ${k(prod.revenue)}k and keeps ${pct(prod.gm_pct, 0)} of it, the best margin in the business, and still loses ${k(-prod.result)}k on ${k(-prod.all_overhead)}k of running costs.`,
      `Consulting sold ${k(cons.revenue)}k against a ${k(cons.budget)}k plan and keeps only ${pct(cons.gm_pct, 0)}, so it still loses ${k(-cons.result)}k.`,
      `The three add to ${k(dTot.result)}k, the same as the five. Nothing is lost in the grouping.`,
    ];
  })(),
  notes: `The same month read three ways instead of five, which is how the Summary tab is actually built.\n\nProduction is the parent of Events and Video. Consulting is the parent of Integration and Consulting. Support has no children. Admin is not a trading department; its costs are spread across the three.\n\n` +
    TRIO.map((x) => `${x.name}: revenue ${money(x.revenue)} against ${money(x.budget)} budget. Gross profit ${money(x.gross_profit)} at ${pct(x.gm_pct)}. All overhead ${money(-x.all_overhead)}. Result ${money(x.result)}.`).join('\n') +
    `\n\nThe three total ${money(dTot.result)}, which is net profit for the month, and the check row on the tab proves it against the five department version above.\n\nWhat the grouping shows that the five did not. Production, Events and Video together, keeps ${pct(TRIO[1].gm_pct)} of what it sells. That is the strongest margin we have anywhere. It loses money only because it carries ${money(-TRIO[1].all_overhead)} of running costs against ${money(TRIO[1].revenue)} of sales. This is a volume problem, not a pricing problem, and it is the clearest growth case in the pack.\n\nConsulting, Integration and Consulting together, is the opposite. It sold ${money(TRIO[2].revenue - TRIO[2].budget)} more than plan and keeps only ${pct(TRIO[2].gm_pct)}, because almost all of that was resold equipment through Integration. Selling more of it will not fix the result.\n\nSupport is the whole profit of the business. On these three lines that is the decision: protect Support, give Production volume, and resolve Consulting.`,
});

/* ---------------------------------- one P&L per department ---------------- */
// The pack's Dept P&L tab, pointed at each department in turn. Same columns
// every time: the month, the same month last year, and budget, each with a
// variance, so the three comparisons a department head asks for are all there.
const DEPT_STORY = {
  Support: (d) => ([
    `Billed ${k(d.revenue)}k against ${k(d.budget)}k. Within ${k(Math.abs(d.revenue - d.budget))}k of plan, which is what fixed monthly contracts should do.`,
    `Keeps ${pct(d.gm_pct, 0)} of what it sells, the same as last September. Steady and predictable.`,
    `After its ${k(-d.all_overhead)}k share of running costs it still made ${k(d.result)}k. The only department that does.`,
    `Of that share, ${money(staff[0].office)} is office wages. Support picks up the most because it works the most hours.`,
    `On the monthly report's basis Support bills ${pct(HU.Support.pct, 0)} of the hours it is available for. The best in the business.`,
  ]),
  Events: (d) => ([
    `Billed ${k(d.revenue)}k against ${k(d.budget)}k, in a month between the busy results periods.`,
    `Keeps ${pct(d.gm_pct, 0)} of what it sells, the second best in the business.`,
    `And still loses ${k(-d.result)}k, because it is charged ${k(-d.all_overhead)}k of running costs on ${k(d.revenue)}k of sales.`,
    `${k(-d.alloc_overhead)}k of that is office cost, split on who uses the office. Events is not the department sitting in it.`,
    `${money(util[1].leave)} hours went to leave and only ${money(util[1].chargeable)} onto jobs. The people were there, the work was not.`,
  ]),
  Video: (d) => ([
    `Billed ${k(d.revenue)}k against ${k(d.budget)}k. Short of plan for the third month running.`,
    `Keeps ${pct(d.gm_pct, 0)} of what it sells, the best margin in the business.`,
    `But on ${k(d.revenue)}k of sales that is only ${k(d.gross_profit)}k, against ${k(-d.all_overhead)}k of running costs.`,
    `The margin is not the problem. There is not enough work going through it.`,
    `${util[2].chargeable.toFixed(1)} billable hours in the month. That is one person, part of the time.`,
  ]),
  Integration: (d) => ([
    `Billed ${k(d.revenue)}k against a ${k(d.budget)}k plan. Nearly all of it one installation job, now finished.`,
    `Keeps only ${pct(d.gm_pct, 0)} of it. We bought the equipment and resold it, so the money went mostly to suppliers.`,
    `${k(d.revenue)}k of sales turned into ${k(d.gross_profit)}k. Do not read the revenue line as a good month.`,
    `It shows a ${k(d.result)}k profit only because it carries almost no running costs, ${k(-d.all_overhead)}k.`,
    `${money(util[3].chargeable)} billable hours against ${k(d.revenue)}k of sales. We resold, we did not deliver.`,
  ]),
  Consulting: (d) => ([
    `Billed ${k(d.revenue)}k against a ${k(d.budget)}k plan. There is almost no fee income here.`,
    `Carries ${k(-d.all_overhead)}k of running costs against ${k(d.revenue)}k of sales, so it loses ${k(-d.result)}k.`,
    `${money(util[4].non_chargeable)} hours went into work we cannot bill against ${money(util[4].chargeable)} we can.`,
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
      `Sales ${money(d.revenue)} against a budget of ${money(d.budget)}, ${d.revenue >= d.budget ? 'ahead' : 'behind'} by ${money(Math.abs(d.revenue - d.budget))}. Last September it billed ${money(d.ly)}.\n\n` +
      `After direct costs of ${money(-d.cos)} it kept ${money(d.gross_profit)}, which is ${pct(d.gm_pct)} of what it sold.\n\n` +
      `Running costs charged to it were ${money(-d.all_overhead)}. ${money(-d.own_overhead)} of that is its own spending and ${money(-d.alloc_overhead)} is its share of the office, charged on the keys set in the pack.\n\n` +
      `That leaves ${money(d.result)} for the month.\n\n` +
      `Year to date it has billed ${money(d.ytd_revenue)} against a budget of ${money(d.ytd_budget)}.\n\n` +
      `On people: employment expense of ${money(staff.find((x) => x.name === name).total)}, of which ${money(staff.find((x) => x.name === name).departmental)} is its own staff and ${money(staff.find((x) => x.name === name).office)} is its share of office wages. It put ${money(util.find((x) => x.name === name).chargeable)} hours onto jobs and ${money(util.find((x) => x.name === name).non_chargeable)} onto work we cannot bill, with ${money(util.find((x) => x.name === name).leave)} hours of leave.`,
  });
}

snipSlide({
  eyebrow: 'Financial performance', title: 'Revenue against budget and last year',
  snip: 'rev_by_dept', source: FROM('Cost Analysis', 'revenue by department, the month'),
  commentary: [
    `Integration billed ${k(depts[3].revenue)}k against a ${k(depts[3].budget)}k plan. Nearly all of it was one big installation job, now finished, so next month drops back.`,
    `Support billed ${k(depts[0].revenue)}k against ${k(depts[0].budget)}k. On plan, which is what fixed monthly contracts should do.`,
    `Events billed ${k(depts[1].revenue)}k against ${k(depts[1].budget)}k, in a month that sits between the busy results periods.`,
    `Video billed ${k(depts[2].revenue)}k against ${k(depts[2].budget)}k and Consulting ${k(depts[4].revenue)}k against ${k(depts[4].budget)}k. Both have been short every month this year.`,
    `Company sales are ${pct(rev.month / rev.month_ly - 1, 0)} up on last September. We are selling more than last year and keeping less of it.`,
  ],
  notes: `Revenue by department for ${MONTH}, actual against budget and against the same month last year.\n\n` +
    byRev.map((d) => `${d.name}: ${money(d.revenue)} against a budget of ${money(d.budget)} and ${money(d.ly)} last year.`).join('\n') +
    `\n\nCompany revenue ${money(rev.month)} against ${money(rev.month_budget)} budget and ${money(rev.month_ly)} last September.\n\nThe beat is Integration, and it is one large fit-out at ${pct(depts[3].gm_pct)} gross margin. Read it as timing, not growth. Video and Consulting are short of budget for the third month running.`,
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
      'Everything below the top three is under six thousand across a full quarter.',
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
    'It rises when our own team runs out of hours, so watch it against the utilisation page.',
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
      `${t[1].name} came from nothing to ${money(t[1].ytd)} in a quarter, and nearly all of it was one job in September.`,
      `After the top five it falls away fast. Number ten is only ${money(t[9].ytd)}.`,
    ];
  })(),
  notes: `Top ten customers by revenue, year to date, with the top three called out in red.\n\n` +
    D.top10.map((x) => `${x.rank}. ${x.name}: ${money(x.ytd)} for the year, ${pct(x.share)} of the total, ${money(x.month)} of it in ${MONTH}.`).join('\n') +
    `\n\nThe ten together are ${money(D.top10_total.top_ten_ytd)}, or ${pct(D.top10_total.top_ten_share)} of everything we sold. Every other customer on the books makes up the remaining ${money(D.top10_total.rest_ytd)}.\n\nThe three red bars are the point of the slide. They are ${pct(D.top10.slice(0, 3).reduce((a, x) => a + x.ytd, 0) / rev.ytd)} of the year between them.\n\nThe difference that matters is contract work against one-off projects. ${D.top10[0].name} is contracted support that renews every month. ${D.top10[1].name} is a fit-out that has finished. Counting them as the same kind of income is how a forecast goes wrong.\n\nWhat I am recommending: a named owner for each of the top three before the half year, and the new business target set on customers four to twenty rather than on the ones we already have.`,
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
      `This is the fastest cash available to the business this quarter. Nothing else comes close.`,
    ];
  })(),
  notes: `How long customers take to pay, at 30 September 2026.\n\nWe are owed ${money(D.dso.receivables)}. Across the last three months we sold ${money(D.dso.revenue_3m)}, which over ${D.dso.days_in || 92} days is ${money(D.dso.receivables / D.dso.days)} of sales a day. Divide what we are owed by that and it takes customers ${D.dso.days.toFixed(1)} days to pay us.\n\nThe target is ${D.dso.target} days. We are ${D.dso.above.toFixed(1)} days over it. If we collected on time we would be owed ${money(D.dso.at_target)} instead of ${money(D.dso.receivables)}, so ${money(D.dso.locked)} of cash that should already be ours is sitting in the debtor book.\n\nPut that next to the profit position. The quarter made ${money(eb.ytd)}. There is ${money(D.dso.locked)} of our own money uncollected. Collecting it is worth many times more than anything we can trade our way to between now and Christmas, and it costs nothing but phone calls.\n\nEvery day we take off the number brings in ${money(D.dso.receivables / D.dso.days)}.`,
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
  notes: `The ten largest amounts owed at 30 September 2026.\n\n` +
    D.debtors.slice(0, 5).map((x) => `${x.name}: ${money(x.amount)}, ${pct(x.share)} of the total.`).join('\n') +
    `\n\nThe ageing: ` + D.ageing.filter((x) => x.amount > 0).map((x) => `${x.bucket} ${money(x.amount)} at ${pct(x.share)}`).join(', ') +
    `. Late in total ${money(D.dso.overdue)}, or ${pct(D.dso.overdue_share)} of the book.\n\nThe action here is narrow and that is the point. Two customers hold most of the balance. A targeted chase on those two is worth more than a general collections exercise across everyone, and it is the fastest money available to us this quarter.`,
});

/* ===================================================== 04 people and utilisation */
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

/* ===================================================== 05 budget and outlook */
section('05', 'Budget and outlook', 'What the rest of the year looks like and what we change',
  'Three months are closed and the remaining nine are still at original budget. That is the decision in front of the board.');

snipSlide({
  eyebrow: 'Budget and outlook', title: 'Revenue by month against budget',
  snip: 'income_month', source: FROM('Cost Analysis', 'total income by month, actual to date then budget'),
  commentary: (() => {
    const m = D.income_by_month;
    const q1a = m.slice(0, 3).reduce((a, x) => a + x.actual, 0), q1b = m.slice(0, 3).reduce((a, x) => a + x.budget, 0);
    return [
      `Q1 came in ${k(q1a - q1b)}k behind on revenue, and almost all of it is Consulting and Video.`,
      'The remaining nine months are untouched original budget. They are not a forecast.',
      `October alone is budgeted at ${k(m[3].budget)}k, nearly double our best month so far.`,
      'December carries a planned loss. The business closes for two weeks, so that is correct and should stay.',
      'Recommend a Q1 rolling forecast: rebuild Consulting on real capacity and reset the margin assumptions.',
    ];
  })(),
  notes: (() => {
    const m = D.income_by_month;
    const q1a = m.slice(0, 3).reduce((a, x) => a + x.actual, 0), q1b = m.slice(0, 3).reduce((a, x) => a + x.budget, 0);
    return `Revenue by month, actual for the three closed months and original budget for the rest.\n\nQ1 actual ${money(q1a)} against a budget of ${money(q1b)}, behind by ${money(q1b - q1a)}.\n\nThe nine months from October are still at the budget set before the year started. October is budgeted at ${money(m[3].budget)}, higher than any month we have delivered.\n\nTwo things for the board. The Q1 variance is concentrated in Consulting and Video, which were budgeted on a resourcing level that does not exist. And the margin assumption in the budget is above what the business has delivered for two years, so even a month that hits the revenue number will miss the profit number.\n\nMy recommendation is a rolling forecast from Q2, rebuilding Consulting around actual capacity and resetting gross margin to delivered rates. December should stay as budgeted: the business closes for two weeks and the loss is real.`;
  })(),
});

/* ===================================================== summary and actions */
{
  const s = pres.addSlide({ masterName: 'CTS_LIGHT' });
  head(s, 'Summary and actions', 'What we agreed');
  const sup = util[0], ev = util[1];
  const gain = (depts[0].revenue / sup.chargeable) * sup.chargeable * 0.05 + (depts[1].revenue / ev.chargeable) * ev.chargeable * 0.05;
  const acts = [
    ['P&L', `Reforecast FY27 on margin, not revenue. At ${pct(ytdMargin)} gross margin break even needs ${k(breakEven)}k a month and we are running at ${k(revRun)}k.`, 'GC and FC', '31 Oct 26'],
    ['Overhead keys', `Events carries ${k(-depts[1].all_overhead)}k of overhead on ${k(depts[1].revenue)}k of revenue and cannot be profitable on that key. Test the keys for FY28, do not change them mid year.`, 'GC and FC', '30 Nov 26'],
    ['Consulting', `${k(-depts[4].all_overhead)}k of running costs against ${k(depts[4].revenue)}k of fees. Decide whether the department is rebuilt, merged into Integration, or closed.`, 'GC', '30 Nov 26'],
    ['Customers', `Top three are ${pct(D.top10.slice(0, 3).reduce((a, x) => a + x.ytd, 0) / rev.ytd)} of revenue. Agree a named retention plan, and set the growth target on the second ten.`, 'GC and Sales', '14 Nov 26'],
    ['Cash', `DSO ${D.dso.days.toFixed(1)} days. Getting to ${D.dso.target} releases ${k(D.dso.locked)}k. Chase ${D.debtors[0].name} and ${D.debtors[1].name} first.`, 'FC', '24 Oct 26'],
    ['Utilisation', `Chargeable per cent is ${pct(HU.Grand.pct, 0)} against the report's own ${pct(HU.Grand.target, 0)} target. Five more points on Support and Production is worth about ${k(gain)}k a month at no extra cost.`, 'Ops and FC', '30 Nov 26'],
    ['Leave', `Production lost ${money(HU.Production.lv)} hours to leave on ${money(HU.Production.workable)} available. Plan leave against the booking calendar, not around it.`, 'Ops', '30 Nov 26'],
  ];
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
  s.addNotes('Seven actions, each with an owner and a date.\n\nThe first two matter most: reforecasting on margin rather than revenue, and testing the overhead keys. Everything else in this pack follows from those two numbers.\n\nThe Consulting decision cannot keep being deferred; it has carried running costs against almost no fees for three months.\n\nThe cash action is the only one that produces money this quarter, and the utilisation action is the largest controllable item.');
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
  s.addNotes(`Close on the two numbers that matter: break even needs ${money(breakEven)} a month at our real margin, and we are running at ${money(revRun)}. Everything else in the pack is detail underneath that.`);
}

module.exports = { pres };
