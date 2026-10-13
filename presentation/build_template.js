/*
 * Builds the CTS PowerPoint template.
 *   node presentation/build_template.js
 *
 * Writes presentation/CTS Presentation Template.pptx (and a .potx copy) with
 * the nineteen layouts defined in cts_masters.js:
 *
 *   Covers (6):   CTS dark, CTS light, CTS 20 years, Production, Consulting, Support
 *   Sections (5): CTS navy, CTS light, Production, Consulting, Support
 *   Content (8):  Header Only and Header and Text, for CTS and each division
 *
 * plus one example slide per layout so the deck opens with something to copy.
 */
const path = require('path');
const cts = require('./cts_masters');

const { C, FONT, DIVISIONS, M } = cts;
const pres = cts.newDeck('CTS Presentation Template');

// ---------------------------------------------------------------------------
// Example slides, one per layout
// ---------------------------------------------------------------------------
function cover(master, eyebrow, title, subtitle, note) {
  const s = pres.addSlide({ masterName: master });
  s.addText(eyebrow, { placeholder: 'eyebrow' });
  s.addText(title, { placeholder: 'title' });
  s.addText(subtitle, { placeholder: 'subtitle' });
  s.addText('Presenter name, role', { placeholder: 'presenter' });
  s.addText('September 2026', { placeholder: 'date' });
  s.addNotes(note);
}
cover('CTS Cover (dark)', 'MONTHLY REPORTING', 'Financial Controller Pack', 'Month-end results, budget variances and the outlook for the quarter', 'CTS Cover (dark). Charcoal background, CTS Seamless AV logo, the three division colours as the motif.');
cover('CTS Cover (light)', 'TRAINING', 'Payroll Processing Walkthrough', 'The fortnightly cycle from payroll notes to bank upload', 'CTS Cover (light). White with a charcoal panel.');
cover('CTS Cover (20 years)', 'COMPANY UPDATE', 'Twenty years of seamless AV', 'Where we have come from and where we are heading', 'CTS Cover (20 years). Uses the 20-year anniversary lock-up currently on ctsav.com.au.');
cover('CTS Production Cover', 'EVENT PROPOSAL', 'Annual General Meeting 2026', 'Hybrid event production, webcast and on-site technical delivery', 'CTS Production cover. Charcoal background with the Production lock-up and mint and teal circles.');
cover('CTS Consulting Cover', 'DESIGN PROPOSAL', 'Collaboration Hub Technology Design', 'Concept, procurement and deployment for the new head office', 'CTS Consulting cover. Charcoal background with the Consulting lock-up and lavender and navy circles.');
cover('CTS Support Cover', 'SERVICE REVIEW', 'Managed Services Quarterly Review', 'Onsite and remote support performance, July to September', 'CTS Support cover. Charcoal background with the Support lock-up and rose and burgundy circles.');

function section(master, number, title, subtitle, note) {
  const s = pres.addSlide({ masterName: master });
  s.addText(number, { placeholder: 'number' });
  s.addText(title, { placeholder: 'title' });
  s.addText(subtitle, { placeholder: 'subtitle' });
  if (note) s.addNotes(note);
}
section('CTS Section (navy)', '01', 'Results for the month', 'Revenue, gross margin and overheads against budget', 'CTS Section (navy). Change the number and title for each section.');
section('CTS Section (light)', '02', 'Outlook and next steps', 'What changes next month and who owns each action', 'CTS Section (light). White version of the divider.');
section('CTS Production Section', '03', 'Event production', 'Delivery: above and beyond', 'Division section dividers carry the division lock-up and colours.');
section('CTS Consulting Section', '04', 'Consulting', 'Strategy: every detail matters');
section('CTS Support Section', '05', 'Managed support', 'Service: always a step ahead');

// CTS Header Only, with a sample native chart in the division colours
{
  const s = pres.addSlide({ masterName: 'CTS Header Only' });
  s.addText('RESULTS FOR THE MONTH', { placeholder: 'section' });
  s.addText('Revenue by division, rolling twelve months', { placeholder: 'title' });
  const labels = ['Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep'];
  s.addChart(pres.ChartType.bar, [
    { name: 'Production', labels, values: [210, 225, 180, 160, 240, 270, 255, 290, 320, 305, 330, 360] },
    { name: 'Consulting', labels, values: [120, 118, 125, 122, 128, 130, 126, 135, 138, 132, 140, 142] },
    { name: 'Support', labels, values: [80, 82, 75, 78, 72, 70, 74, 65, 62, 68, 60, 58] },
  ], {
    x: M, y: 1.7, w: 8.9, h: 3.3, barDir: 'col', barGrouping: 'stacked', chartColors: [C.MINT, C.LAVENDER, C.ROSE],
    showLegend: true, legendPos: 'b', legendFontFace: FONT, legendFontSize: 10, legendColor: C.GREY,
    showValue: false,
    catAxisLabelColor: C.GREY, catAxisLabelFontSize: 10, catAxisLabelFontFace: FONT, catGridLine: { style: 'none' },
    valAxisLabelColor: C.GREY, valAxisLabelFontSize: 10, valAxisLabelFontFace: FONT, valGridLine: { color: C.LINE, size: 0.5 },
    valAxisLineShow: false, catAxisLineShow: false,
  });
  s.addNotes('CTS Header Only. Delete the sample chart and drop in your own chart, table or image. Chart values are illustrative only. Division colours: Production mint, Consulting lavender, Support rose.');
}
// CTS Header and Text, bullets
{
  const s = pres.addSlide({ masterName: 'CTS Header and Text' });
  s.addText('RESULTS FOR THE MONTH', { placeholder: 'section' });
  s.addText('Three things to take from this month', { placeholder: 'title' });
  s.addText([
    { text: 'Revenue finished ahead of budget, driven by the top five clients.', options: { bullet: true, breakLine: true } },
    { text: 'Gross margin held steady despite higher casual labour hours.', options: { bullet: true, breakLine: true } },
    { text: 'Overheads were under budget, with the largest saving in software subscriptions.', options: { bullet: true, breakLine: true } },
    { text: 'Cash position improved as aged receivables were collected.', options: { bullet: true } },
  ], { placeholder: 'body' });
  s.addNotes('CTS Header and Text, bullet version. Keep to four or five bullets per slide.');
}
// CTS Header and Text, callout card
{
  const s = pres.addSlide({ masterName: 'CTS Header and Text' });
  s.addText('OUTLOOK AND NEXT STEPS', { placeholder: 'section' });
  s.addText('Header and text with a callout card', { placeholder: 'title' });
  s.addText([
    { text: 'Body text goes here. Keep paragraphs short and lead with the point you want the audience to remember.', options: { bullet: false, breakLine: true } },
    { text: '', options: { bullet: false, breakLine: true } },
    { text: 'First supporting point', options: { bullet: { indent: 18 }, breakLine: true } },
    { text: 'Second supporting point', options: { bullet: { indent: 18 }, breakLine: true } },
    { text: 'Third supporting point', options: { bullet: { indent: 18 } } },
  ], { x: M, y: 1.7, w: 4.9, h: 3.3, margin: 0, isTextBox: true, valign: 'top', fontFace: FONT, fontSize: 16, color: C.BLACK, paraSpaceAfter: 8, bullet: false });
  s.addShape(pres.ShapeType.rect, { x: 5.8, y: 1.7, w: 3.65, h: 3.3, fill: { color: C.MINT_TINT }, line: { color: C.MINT_TINT, width: 0 } });
  s.addText('KEY POINT', { x: 6.1, y: 2.05, w: 3.05, h: 0.3, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 10, bold: true, color: C.TEAL, charSpacing: 2 });
  s.addText('A callout, figure or quote that supports the text on the left', { x: 6.1, y: 2.4, w: 3.05, h: 1.6, margin: 0, isTextBox: true, valign: 'top', fontFace: FONT, fontSize: 18, bold: true, color: C.BLACK });
  s.addText('Optional source or note', { x: 6.1, y: 4.35, w: 3.05, h: 0.3, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 11, color: C.GREY });
  s.addNotes('CTS Header and Text with a narrower text box on the left and a callout card on the right.');
}
// Division content examples
const DIVISION_EXAMPLES = {
  production: { section: 'EVENT PRODUCTION', headerOnly: 'Run sheet and technical plan', bullets: ['Hybrid AGM with in-room audience and secure webcast.', 'Three-camera vision mix with lower thirds and live polling.', 'Rehearsal the afternoon before, full technical check on the day.', 'Post-event edit delivered within two business days.'] },
  consulting: { section: 'AUDIO VISUAL DESIGN', headerOnly: 'Proposed room standards', bullets: ['Vendor-agnostic design across meeting rooms, boardroom and town hall space.', 'One-touch join on every room with a single control standard.', 'Procurement, deployment and commissioning managed end to end.', 'Ongoing optimisation reviewed each quarter.'] },
  support: { section: 'MANAGED SERVICES', headerOnly: 'Ticket volume and response times', bullets: ['Onsite team covering business hours with remote backup after hours.', 'All priority-one tickets responded to within the agreed window.', 'Back-up staff guarantee applied across leave and illness.', 'Preventative maintenance completed on every room this quarter.'] },
};
for (const d of DIVISIONS) {
  const ex = DIVISION_EXAMPLES[d.key];
  {
    const s = pres.addSlide({ masterName: `CTS ${d.name} Header Only` });
    s.addText(ex.section, { placeholder: 'section' });
    s.addText(ex.headerOnly, { placeholder: 'title' });
    s.addNotes(`CTS ${d.name} Header Only. Empty body for a chart, table, plan or image.`);
  }
  {
    const s = pres.addSlide({ masterName: `CTS ${d.name} Header and Text` });
    s.addText(ex.section, { placeholder: 'section' });
    s.addText(`What CTS ${d.name} delivers`, { placeholder: 'title' });
    s.addText(ex.bullets.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < ex.bullets.length - 1 } })), { placeholder: 'body' });
  }
}

cts.writeDeck(pres, path.join(__dirname, 'CTS Presentation Template.pptx'), { potx: true })
  .catch((err) => { console.error(err); process.exit(1); });
