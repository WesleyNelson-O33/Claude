/*
 * Builds the CTS PowerPoint template.
 *   node presentation/build_template.js
 *
 * Writes presentation/CTS Presentation Template.pptx (and a .potx copy) with
 * nineteen layouts built into the Slide Master:
 *
 *   Covers (6):   CTS dark, CTS light, CTS 20 years, Production, Consulting, Support
 *   Sections (5): CTS navy, CTS light, Production, Consulting, Support
 *   Content (8):  Header Only and Header and Text, for CTS and each division
 *
 * plus one example slide per layout so the deck opens with something to copy.
 *
 * Palette: the official CTS Seamless AV brand palette (brand-source/
 * cts_pantonevalues.pdf). Base black plus three division colours, each with a
 * light and a dark variant. Greys and the mint tint are derived here for
 * footers and cards; they are not brand colours.
 *
 * Logos: presentation/assets (see assets/README.md for sources).
 */
const path = require('path');
const fs = require('fs');
const pptxgen = require('pptxgenjs');
const JSZip = require('jszip');

// Brand palette (CTS Seamless AV)
const BLACK = '121820';    // Base black, PMS Black 6 C
const MINT = '3FD0C9';     // Production (light), PMS 3255 C
const LAVENDER = '6D71FF'; // Consulting (light), PMS 2124 C
const ROSE = 'F33844';     // Support (light), PMS 1788 C
const TEAL = '005358';     // Production (dark), PMS 7476 C
const NAVY = '122B82';     // Consulting (dark), PMS 287 C
const BURGUNDY = 'A8052E'; // Support (dark), PMS 1945 C
const WHITE = 'FFFFFF';

// Derived neutrals (not brand colours)
const GREY = '5F6670';     // captions, footers, subtitles on white
const SOFT = 'C9CED6';     // secondary text on black or navy
const MINT_TINT = 'E8F9F8'; // callout cards on white
const LINE = 'E3E6EA';     // chart gridlines

const FONT = 'Calibri'; // Brand font is Tomato Grotesk; Calibri is the safe fallback that ships with Office.

// The three divisions: light colour for fills on dark, dark colour for text on white.
const DIVISIONS = [
  { key: 'production', name: 'Production', light: MINT, dark: TEAL, promise: 'Engaging Experiences' },
  { key: 'consulting', name: 'Consulting', light: LAVENDER, dark: NAVY, promise: 'End to End collaboration' },
  { key: 'support', name: 'Support', light: ROSE, dark: BURGUNDY, promise: 'Personalised Service' },
];

// ---------------------------------------------------------------------------
// Logo files
// ---------------------------------------------------------------------------
const ASSETS = path.join(__dirname, 'assets');
function logoFile(kind, onDark) {
  // kind: 'cts' | '20yrs' | 'production' | 'consulting' | 'support'
  // onDark: false = charcoal version for white; true = white version for black;
  //         'white' = all-white version for a division-coloured background
  const base = kind === 'cts' ? 'cts-logo' : kind === '20yrs' ? 'cts-logo-20yrs' : `cts-${kind}`;
  const names = onDark === 'white' ? [`${base}-white.png`, `${base}-dark.png`] : onDark ? [`${base}-dark.png`, `${base}.png`] : [`${base}.png`];
  for (const n of names) {
    const f = path.join(ASSETS, n);
    if (fs.existsSync(f)) return f;
  }
  throw new Error(`Missing logo file for ${kind} (${onDark ? 'dark' : 'light'}) in presentation/assets`);
}
function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}
// Place a logo by height. alignRight anchors the image's right edge at x.
function logo(objects, kind, { x, y, h, onDark, alignRight }) {
  const file = logoFile(kind, onDark);
  const px = pngSize(file);
  const w = h * (px.w / px.h);
  objects.push({ image: { path: file, x: alignRight ? x - w : x, y, w, h } });
}

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // 10in x 5.625in
pres.author = 'Corporate Technology Services';
pres.company = 'Corporate Technology Services';
pres.title = 'CTS Presentation Template';

const W = 10;
const H = 5.625;
const M = 0.55; // side margin

// ---------------------------------------------------------------------------
// Shared pieces
// ---------------------------------------------------------------------------
function circle(objects, { x, y, d, color, transparency }) {
  objects.push({ rect: { x, y, w: d, h: d, fill: { color, transparency }, line: { color, width: 0, transparency: 100 } } });
}
// The three division colours overlapping: the cover motif.
function trio(objects, ox, oy, scale) {
  circle(objects, { x: ox, y: oy, d: 2.6 * scale, color: MINT, transparency: 15 });
  circle(objects, { x: ox + 1.35 * scale, y: oy + 1.0 * scale, d: 2.6 * scale, color: LAVENDER, transparency: 25 });
  circle(objects, { x: ox + 0.35 * scale, y: oy + 2.1 * scale, d: 2.6 * scale, color: ROSE, transparency: 30 });
}
function footer(objects, onDark, label) {
  objects.push({
    text: {
      text: label,
      options: { x: M, y: H - 0.42, w: 5, h: 0.25, margin: 0, valign: 'middle', fontFace: FONT, fontSize: 9, color: onDark ? SOFT : GREY },
    },
  });
}
function slideNumber(onDark) {
  return { x: W - M - 0.6, y: H - 0.42, w: 0.6, h: 0.25, fontFace: FONT, fontSize: 9, color: onDark ? SOFT : GREY, align: 'right' };
}
function ph(name, type, opts, text) {
  return { placeholder: { options: { name, type, margin: 0, fontFace: FONT, ...opts }, text } };
}

// ---------------------------------------------------------------------------
// Cover layouts
// ---------------------------------------------------------------------------
// Shared text placeholders for a cover: eyebrow, title, subtitle, presenter, date.
function coverText(objects, { textW, eyebrowColor, titleColor, subColor, metaColor, titleSize = 40, dateX = 6.5, dateW = 2.95 }) {
  objects.push(ph('eyebrow', 'body', { x: M, y: 2.0, w: textW, h: 0.35, fontSize: 12, bold: true, color: eyebrowColor, charSpacing: 3 }, 'PRESENTATION TYPE'));
  objects.push(ph('title', 'title', { align: 'left', x: M, y: 2.35, w: textW, h: 1.35, valign: 'top', fontSize: titleSize, bold: true, color: titleColor }, 'Presentation title'));
  objects.push(ph('subtitle', 'body', { x: M, y: 3.7, w: textW, h: 0.6, valign: 'top', fontSize: 18, color: subColor }, 'Subtitle or one-line description'));
  objects.push(ph('presenter', 'body', { x: M, y: 4.65, w: 4.0, h: 0.55, valign: 'bottom', fontSize: 12, color: metaColor }, 'Presenter name, role'));
  objects.push(ph('date', 'body', { x: dateX, y: 4.65, w: dateW, h: 0.55, valign: 'bottom', align: 'right', fontSize: 12, color: metaColor }, 'Month Year'));
}

// 1. CTS Cover (dark)
{
  const objects = [];
  trio(objects, 6.9, -0.6, 1.45);
  logo(objects, 'cts', { x: M, y: 0.45, h: 0.85, onDark: true });
  coverText(objects, { textW: 6.2, eyebrowColor: MINT, titleColor: WHITE, subColor: SOFT, metaColor: SOFT });
  pres.defineSlideMaster({ title: 'CTS Cover (dark)', background: { color: BLACK }, objects });
}
// 2. CTS Cover (light)
{
  const objects = [];
  objects.push({ rect: { x: 6.9, y: 0, w: 3.1, h: H, fill: { color: BLACK }, line: { color: BLACK, width: 0 } } });
  trio(objects, 6.5, 0.4, 1.25);
  objects.push({ rect: { x: 0, y: 0, w: 6.9, h: H, fill: { color: WHITE }, line: { color: WHITE, width: 0 } } }); // mask
  logo(objects, 'cts', { x: M, y: 0.45, h: 0.85, onDark: false });
  coverText(objects, { textW: 5.9, eyebrowColor: NAVY, titleColor: BLACK, subColor: GREY, metaColor: GREY, titleSize: 36, dateX: 4.2, dateW: 2.4 });
  pres.defineSlideMaster({ title: 'CTS Cover (light)', background: { color: WHITE }, objects });
}
// 3. CTS Cover (20 years)
{
  const objects = [];
  circle(objects, { x: 6.6, y: 2.2, d: 5.2, color: NAVY, transparency: 55 });
  logo(objects, '20yrs', { x: M, y: 0.5, h: 1.05, onDark: true });
  coverText(objects, { textW: 6.4, eyebrowColor: MINT, titleColor: WHITE, subColor: SOFT, metaColor: SOFT });
  pres.defineSlideMaster({ title: 'CTS Cover (20 years)', background: { color: BLACK }, objects });
}
// 4 to 6. Division covers
for (const d of DIVISIONS) {
  const objects = [];
  circle(objects, { x: 6.4, y: -1.6, d: 4.6, color: d.light, transparency: 10 });
  circle(objects, { x: 7.6, y: 2.4, d: 4.2, color: d.dark, transparency: 20 });
  logo(objects, d.key, { x: M, y: 0.45, h: 0.85, onDark: true });
  coverText(objects, { textW: 6.0, eyebrowColor: d.light, titleColor: WHITE, subColor: SOFT, metaColor: SOFT });
  pres.defineSlideMaster({ title: `CTS ${d.name} Cover`, background: { color: BLACK }, objects });
}

// ---------------------------------------------------------------------------
// Section layouts
// ---------------------------------------------------------------------------
function sectionBody(objects, { blockColor, numberColor, titleColor, subColor, labelColor, onDark, footerLabel }) {
  objects.push({ text: { text: 'SECTION', options: { x: M, y: 0.45, w: 3, h: 0.35, margin: 0, valign: 'middle', fontFace: FONT, fontSize: 11, bold: true, color: labelColor, charSpacing: 3 } } });
  objects.push({ rect: { x: M, y: 2.05, w: 1.4, h: 1.4, fill: { color: blockColor }, line: { color: blockColor, width: 0 } } });
  objects.push(ph('number', 'body', { x: M, y: 2.05, w: 1.4, h: 1.4, align: 'center', valign: 'middle', fontSize: 48, bold: true, color: numberColor }, '01'));
  objects.push(ph('title', 'title', { align: 'left', x: M + 1.75, y: 2.05, w: 6.6, h: 0.9, valign: 'bottom', fontSize: 36, bold: true, color: titleColor }, 'Section title'));
  objects.push(ph('subtitle', 'body', { x: M + 1.75, y: 2.98, w: 6.6, h: 0.5, valign: 'top', fontSize: 16, color: subColor }, 'One line describing what this section covers'));
  footer(objects, onDark, footerLabel);
}
// 7. CTS Section (navy)
{
  const objects = [];
  circle(objects, { x: 7.0, y: 2.4, d: 4.6, color: LAVENDER, transparency: 55 });
  logo(objects, 'cts', { x: W - M, y: 0.42, h: 0.6, onDark: true, alignRight: true });
  sectionBody(objects, { blockColor: MINT, numberColor: BLACK, titleColor: WHITE, subColor: SOFT, labelColor: MINT, onDark: true, footerLabel: 'Corporate Technology Services  |  Seamless AV' });
  pres.defineSlideMaster({ title: 'CTS Section (navy)', background: { color: NAVY }, objects, slideNumber: slideNumber(true) });
}
// 8. CTS Section (light)
{
  const objects = [];
  circle(objects, { x: 7.0, y: 2.4, d: 4.6, color: MINT, transparency: 80 });
  logo(objects, 'cts', { x: W - M, y: 0.42, h: 0.6, onDark: false, alignRight: true });
  sectionBody(objects, { blockColor: BLACK, numberColor: WHITE, titleColor: BLACK, subColor: GREY, labelColor: NAVY, onDark: false, footerLabel: 'Corporate Technology Services  |  Seamless AV' });
  pres.defineSlideMaster({ title: 'CTS Section (light)', background: { color: WHITE }, objects, slideNumber: slideNumber(false) });
}
// 9 to 11. Division sections
for (const d of DIVISIONS) {
  const objects = [];
  circle(objects, { x: 7.0, y: 2.4, d: 4.6, color: d.light, transparency: 60 });
  logo(objects, d.key, { x: W - M, y: 0.42, h: 0.6, onDark: 'white', alignRight: true });
  sectionBody(objects, { blockColor: d.light, numberColor: BLACK, titleColor: WHITE, subColor: SOFT, labelColor: d.light, onDark: true, footerLabel: `CTS ${d.name}  |  ${d.promise}` });
  pres.defineSlideMaster({ title: `CTS ${d.name} Section`, background: { color: d.dark }, objects, slideNumber: slideNumber(true) });
}

// ---------------------------------------------------------------------------
// Content layouts
// ---------------------------------------------------------------------------
function contentHeader(objects, { logoKind, labelColor, footerLabel }) {
  objects.push(ph('section', 'body', { x: M, y: 0.4, w: 7.0, h: 0.3, valign: 'middle', fontSize: 10, bold: true, color: labelColor, charSpacing: 3 }, 'SECTION NAME'));
  objects.push(ph('title', 'title', { align: 'left', x: M, y: 0.7, w: 7.6, h: 0.75, valign: 'middle', fontSize: 28, bold: true, color: BLACK }, 'Slide heading'));
  logo(objects, logoKind, { x: W - M, y: 0.42, h: 0.6, onDark: false, alignRight: true });
  footer(objects, false, footerLabel);
}
function bodyPlaceholder(objects) {
  objects.push(ph('body', 'body', { x: M, y: 1.7, w: 8.9, h: 3.3, valign: 'top', fontSize: 16, color: BLACK, paraSpaceAfter: 8, bullet: { indent: 18 } }, 'Body text'));
}
const CONTENT = [
  { key: 'cts', prefix: 'CTS', labelColor: NAVY, footerLabel: 'Corporate Technology Services  |  Seamless AV' },
  ...DIVISIONS.map((d) => ({ key: d.key, prefix: `CTS ${d.name}`, labelColor: d.dark, footerLabel: `CTS ${d.name}  |  ${d.promise}` })),
];
for (const c of CONTENT) {
  {
    const objects = [];
    contentHeader(objects, { logoKind: c.key, labelColor: c.labelColor, footerLabel: c.footerLabel });
    pres.defineSlideMaster({ title: `${c.prefix} Header Only`, background: { color: WHITE }, objects, slideNumber: slideNumber(false) });
  }
  {
    const objects = [];
    contentHeader(objects, { logoKind: c.key, labelColor: c.labelColor, footerLabel: c.footerLabel });
    bodyPlaceholder(objects);
    pres.defineSlideMaster({ title: `${c.prefix} Header and Text`, background: { color: WHITE }, objects, slideNumber: slideNumber(false) });
  }
}

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
cover('CTS Cover (dark)', 'MONTHLY REPORTING', 'Financial Controller Pack', 'Month-end results, budget variances and the outlook for the quarter', 'CTS Cover (dark). Black background, CTS Seamless AV logo, the three division colours as the motif.');
cover('CTS Cover (light)', 'TRAINING', 'Payroll Processing Walkthrough', 'The fortnightly cycle from payroll notes to bank upload', 'CTS Cover (light). White with a black panel.');
cover('CTS Cover (20 years)', 'COMPANY UPDATE', 'Twenty years of seamless AV', 'Where we have come from and where we are heading', 'CTS Cover (20 years). Uses the 20-year anniversary lock-up currently on ctsav.com.au.');
cover('CTS Production Cover', 'EVENT PROPOSAL', 'Annual General Meeting 2026', 'Hybrid event production, webcast and on-site technical delivery', 'CTS Production cover. Black background with the Production lock-up and mint and teal circles.');
cover('CTS Consulting Cover', 'DESIGN PROPOSAL', 'Collaboration Hub Technology Design', 'Concept, procurement and deployment for the new head office', 'CTS Consulting cover. Black background with the Consulting lock-up and lavender and navy circles.');
cover('CTS Support Cover', 'SERVICE REVIEW', 'Managed Services Quarterly Review', 'Onsite and remote support performance, July to September', 'CTS Support cover. Black background with the Support lock-up and rose and burgundy circles.');

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
    x: M, y: 1.7, w: 8.9, h: 3.3, barDir: 'col', barGrouping: 'stacked', chartColors: [MINT, LAVENDER, ROSE],
    showLegend: true, legendPos: 'b', legendFontFace: FONT, legendFontSize: 10, legendColor: GREY,
    showValue: false,
    catAxisLabelColor: GREY, catAxisLabelFontSize: 10, catAxisLabelFontFace: FONT, catGridLine: { style: 'none' },
    valAxisLabelColor: GREY, valAxisLabelFontSize: 10, valAxisLabelFontFace: FONT, valGridLine: { color: LINE, size: 0.5 },
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
  ], { x: M, y: 1.7, w: 4.9, h: 3.3, margin: 0, isTextBox: true, valign: 'top', fontFace: FONT, fontSize: 16, color: BLACK, paraSpaceAfter: 8, bullet: false });
  s.addShape(pres.ShapeType.rect, { x: 5.8, y: 1.7, w: 3.65, h: 3.3, fill: { color: MINT_TINT }, line: { color: MINT_TINT, width: 0 } });
  s.addText('KEY POINT', { x: 6.1, y: 2.05, w: 3.05, h: 0.3, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 10, bold: true, color: TEAL, charSpacing: 2 });
  s.addText('A callout, figure or quote that supports the text on the left', { x: 6.1, y: 2.4, w: 3.05, h: 1.6, margin: 0, isTextBox: true, valign: 'top', fontFace: FONT, fontSize: 18, bold: true, color: BLACK });
  s.addText('Optional source or note', { x: 6.1, y: 4.35, w: 3.05, h: 0.3, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 11, color: GREY });
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

// ---------------------------------------------------------------------------
// Post-process the package.
//  - pptxgenjs draws every master `rect` as a rectangle; the decorative
//    circles are the only master shapes with a transparent fill, so turn
//    those into ellipses.
//  - Also save a .potx copy (same package, template content type) so it can
//    be installed as a PowerPoint template.
// ---------------------------------------------------------------------------
async function finish() {
  const out = path.join(__dirname, 'CTS Presentation Template.pptx');
  await pres.writeFile({ fileName: out });
  const zip = await JSZip.loadAsync(fs.readFileSync(out));
  for (const name of Object.keys(zip.files)) {
    if (!/^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(name)) continue;
    let xml = await zip.file(name).async('string');
    xml = xml.replace(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g, (sp) =>
      sp.includes('<a:alpha ') ? sp.replace('prst="rect"', 'prst="ellipse"') : sp);
    zip.file(name, xml);
  }
  fs.writeFileSync(out, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log('wrote', out);

  const ct = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', ct.replace(
    'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'));
  const potx = out.replace(/\.pptx$/, '.potx');
  fs.writeFileSync(potx, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log('wrote', potx);
}

finish().catch((err) => { console.error(err); process.exit(1); });
