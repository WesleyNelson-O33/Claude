/*
 * Builds the CTS PowerPoint template.
 *   node presentation/build_template.js
 *
 * Writes presentation/CTS Presentation Template.pptx (and a .potx copy) with
 * five slide layouts built into the Slide Master (two covers, a section
 * divider, header only, header and text) and one example slide per layout so
 * the deck opens with something to copy.
 *
 * Palette: the official CTS Seamless AV brand palette, from the colour
 * specification published with the CTS brand guidelines
 * (brandpad.io/cts-seamless, file cts_pantonevalues.pdf). Base black plus
 * three division colours, each with a light and a dark variant. The greys
 * and the mint tint are derived here for footers and cards; they are not
 * brand colours.
 */
const path = require('path');
const fs = require('fs');
const pptxgen = require('pptxgenjs');
const JSZip = require('jszip');

// Brand palette (CTS Seamless AV)
const BLACK = '121820';    // Base black
const MINT = '3FD0C9';     // Production (light), PMS 3255 C
const LAVENDER = '6D71FF'; // Consulting (light), PMS 2124 C
const ROSE = 'F33844';     // Support (light), PMS 1788 C
const TEAL = '005358';     // Production (dark), PMS 7476 C
const NAVY = '122B82';     // Consulting (dark), PMS 287 C
const BURGUNDY = 'A8052E'; // Support (dark)
const WHITE = 'FFFFFF';

// Derived neutrals (not brand colours)
const GREY = '5F6670';     // captions, footers, subtitles on white
const SOFT = 'C9CED6';     // secondary text on black or navy
const MINT_TINT = 'E8F9F8'; // callout cards on white
const LINE = 'E3E6EA';     // chart gridlines

const FONT = 'Calibri'; // Brand font is Tomato Grotesk; Calibri is the safe fallback that ships with Office.

// Real logo files, when present, replace the placeholder mark on every layout.
//   presentation/assets/cts-logo.png       for white backgrounds
//   presentation/assets/cts-logo-dark.png  for black or navy backgrounds (falls back to cts-logo.png)
// Any PNG or JPG works; the file is scaled to the logo height, so supply the
// highest resolution you have. Aspect ratio is read from the file.
const ASSETS = path.join(__dirname, 'assets');
function findLogo(onDark) {
  const names = onDark ? ['cts-logo-dark.png', 'cts-logo-dark.jpg', 'cts-logo.png', 'cts-logo.jpg'] : ['cts-logo.png', 'cts-logo.jpg'];
  for (const n of names) {
    const f = path.join(ASSETS, n);
    if (fs.existsSync(f)) return f;
  }
  return null;
}
function pngSize(file) {
  const b = fs.readFileSync(file);
  if (b.slice(1, 4).toString() === 'PNG') return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
  // JPEG: walk the markers to the first SOF
  let i = 2;
  while (i < b.length) {
    if (b[i] !== 0xff) { i++; continue; }
    const marker = b[i + 1];
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return { h: b.readUInt16BE(i + 5), w: b.readUInt16BE(i + 7) };
    }
    i += 2 + b.readUInt16BE(i + 2);
  }
  return { w: 1, h: 1 };
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

// Placeholder logo mark: square with "CTS" plus the brand lock-up.
// Swap these for the real logo image on the Slide Master once you have the file.
function logo(objects, { x, y, size, onDark, withName }) {
  const file = findLogo(onDark);
  if (file) {
    const px = pngSize(file);
    const h = withName ? size : size;
    const w = h * (px.w / px.h);
    objects.push({ image: { path: file, x, y, w, h } });
    return;
  }
  const box = onDark ? WHITE : BLACK;
  const ink = onDark ? BLACK : WHITE;
  objects.push({ rect: { x, y, w: size, h: size, fill: { color: box }, line: { color: box, width: 0 } } });
  objects.push({
    text: {
      text: 'CTS',
      options: { x, y, w: size, h: size, align: 'center', valign: 'middle', margin: 0, fontFace: FONT, fontSize: size * 26, bold: true, color: ink, charSpacing: 1 },
    },
  });
  if (withName) {
    objects.push({
      text: {
        text: 'Corporate Technology Services',
        options: { x: x + size + 0.15, y: y - 0.02, w: 4.2, h: size * 0.55, valign: 'bottom', margin: 0, fontFace: FONT, fontSize: 12, bold: true, color: onDark ? WHITE : BLACK },
      },
    });
    objects.push({
      text: {
        text: 'Seamless AV',
        options: { x: x + size + 0.15, y: y + size * 0.53, w: 4.2, h: size * 0.47, valign: 'top', margin: 0, fontFace: FONT, fontSize: 10, color: onDark ? SOFT : GREY },
      },
    });
  }
}

// Circles used as the cover motif: the three division colours overlapping.
function circle(objects, { x, y, d, color, transparency }) {
  objects.push({ rect: { x, y, w: d, h: d, fill: { color, transparency }, line: { color, width: 0, transparency: 100 } } });
}

function trio(objects, ox, oy, scale) {
  circle(objects, { x: ox, y: oy, d: 2.6 * scale, color: MINT, transparency: 15 });
  circle(objects, { x: ox + 1.35 * scale, y: oy + 1.0 * scale, d: 2.6 * scale, color: LAVENDER, transparency: 25 });
  circle(objects, { x: ox + 0.35 * scale, y: oy + 2.1 * scale, d: 2.6 * scale, color: ROSE, transparency: 30 });
}

function footer(objects, onDark) {
  objects.push({
    text: {
      text: 'Corporate Technology Services  |  Seamless AV',
      options: { x: M, y: H - 0.42, w: 4.5, h: 0.25, margin: 0, valign: 'middle', fontFace: FONT, fontSize: 9, color: onDark ? SOFT : GREY },
    },
  });
}

function slideNumber(onDark) {
  return { x: W - M - 0.6, y: H - 0.42, w: 0.6, h: 0.25, fontFace: FONT, fontSize: 9, color: onDark ? SOFT : GREY, align: 'right' };
}

// ---------------------------------------------------------------------------
// Layout 1: Cover (dark)
// ---------------------------------------------------------------------------
{
  const objects = [];
  trio(objects, 6.9, -0.6, 1.45);
  logo(objects, { x: M, y: 0.45, size: 0.5, onDark: true, withName: true });
  objects.push({ placeholder: { options: { name: 'eyebrow', type: 'body', x: M, y: 2.0, w: 6.2, h: 0.35, margin: 0, fontFace: FONT, fontSize: 12, bold: true, color: MINT, charSpacing: 3 }, text: 'PRESENTATION TYPE' } });
  objects.push({ placeholder: { options: { name: 'title', type: 'title', align: 'left', x: M, y: 2.35, w: 6.2, h: 1.35, margin: 0, valign: 'top', fontFace: FONT, fontSize: 40, bold: true, color: WHITE }, text: 'Presentation title' } });
  objects.push({ placeholder: { options: { name: 'subtitle', type: 'body', x: M, y: 3.7, w: 6.2, h: 0.6, margin: 0, valign: 'top', fontFace: FONT, fontSize: 18, color: SOFT }, text: 'Subtitle or one-line description' } });
  objects.push({ placeholder: { options: { name: 'presenter', type: 'body', x: M, y: 4.65, w: 4.0, h: 0.55, margin: 0, valign: 'bottom', fontFace: FONT, fontSize: 12, color: SOFT }, text: 'Presenter name, role' } });
  objects.push({ placeholder: { options: { name: 'date', type: 'body', x: 6.5, y: 4.65, w: 2.95, h: 0.55, margin: 0, valign: 'bottom', align: 'right', fontFace: FONT, fontSize: 12, color: SOFT }, text: 'Month Year' } });
  pres.defineSlideMaster({ title: 'CTS Cover (dark)', background: { color: BLACK }, objects });
}

// ---------------------------------------------------------------------------
// Layout 2: Cover (light)
// ---------------------------------------------------------------------------
{
  const objects = [];
  objects.push({ rect: { x: 6.9, y: 0, w: 3.1, h: H, fill: { color: BLACK }, line: { color: BLACK, width: 0 } } });
  trio(objects, 6.5, 0.4, 1.25);
  // white mask so the circles stay inside the black panel
  objects.push({ rect: { x: 0, y: 0, w: 6.9, h: H, fill: { color: WHITE }, line: { color: WHITE, width: 0 } } });
  logo(objects, { x: M, y: 0.45, size: 0.5, onDark: false, withName: true });
  objects.push({ placeholder: { options: { name: 'eyebrow', type: 'body', x: M, y: 2.0, w: 5.8, h: 0.35, margin: 0, fontFace: FONT, fontSize: 12, bold: true, color: NAVY, charSpacing: 3 }, text: 'PRESENTATION TYPE' } });
  objects.push({ placeholder: { options: { name: 'title', type: 'title', align: 'left', x: M, y: 2.35, w: 5.9, h: 1.35, margin: 0, valign: 'top', fontFace: FONT, fontSize: 36, bold: true, color: BLACK }, text: 'Presentation title' } });
  objects.push({ placeholder: { options: { name: 'subtitle', type: 'body', x: M, y: 3.7, w: 5.9, h: 0.6, margin: 0, valign: 'top', fontFace: FONT, fontSize: 16, color: GREY }, text: 'Subtitle or one-line description' } });
  objects.push({ placeholder: { options: { name: 'presenter', type: 'body', x: M, y: 4.65, w: 3.4, h: 0.55, margin: 0, valign: 'bottom', fontFace: FONT, fontSize: 12, color: GREY }, text: 'Presenter name, role' } });
  objects.push({ placeholder: { options: { name: 'date', type: 'body', x: 4.2, y: 4.65, w: 2.4, h: 0.55, margin: 0, valign: 'bottom', align: 'right', fontFace: FONT, fontSize: 12, color: GREY }, text: 'Month Year' } });
  pres.defineSlideMaster({ title: 'CTS Cover (light)', background: { color: WHITE }, objects });
}

// ---------------------------------------------------------------------------
// Layout 3: Section divider (navy)
// ---------------------------------------------------------------------------
{
  const objects = [];
  circle(objects, { x: 7.0, y: 2.4, d: 4.6, color: LAVENDER, transparency: 55 });
  objects.push({ text: { text: 'SECTION', options: { x: M, y: 0.45, w: 3, h: 0.35, margin: 0, valign: 'middle', fontFace: FONT, fontSize: 11, bold: true, color: MINT, charSpacing: 3 } } });
  logo(objects, { x: W - M - 0.42, y: 0.42, size: 0.42, onDark: true, withName: false });
  objects.push({ rect: { x: M, y: 2.05, w: 1.4, h: 1.4, fill: { color: MINT }, line: { color: MINT, width: 0 } } });
  objects.push({ placeholder: { options: { name: 'number', type: 'body', x: M, y: 2.05, w: 1.4, h: 1.4, margin: 0, align: 'center', valign: 'middle', fontFace: FONT, fontSize: 48, bold: true, color: BLACK }, text: '01' } });
  objects.push({ placeholder: { options: { name: 'title', type: 'title', align: 'left', x: M + 1.75, y: 2.05, w: 6.6, h: 0.9, margin: 0, valign: 'bottom', fontFace: FONT, fontSize: 36, bold: true, color: WHITE }, text: 'Section title' } });
  objects.push({ placeholder: { options: { name: 'subtitle', type: 'body', x: M + 1.75, y: 2.98, w: 6.6, h: 0.5, margin: 0, valign: 'top', fontFace: FONT, fontSize: 16, color: SOFT }, text: 'One line describing what this section covers' } });
  footer(objects, true);
  pres.defineSlideMaster({ title: 'CTS Section', background: { color: NAVY }, objects, slideNumber: slideNumber(true) });
}

// ---------------------------------------------------------------------------
// Shared header for the two content layouts
// ---------------------------------------------------------------------------
function contentHeader(objects) {
  objects.push({ placeholder: { options: { name: 'section', type: 'body', x: M, y: 0.4, w: 7.5, h: 0.3, margin: 0, valign: 'middle', fontFace: FONT, fontSize: 10, bold: true, color: NAVY, charSpacing: 3 }, text: 'SECTION NAME' } });
  objects.push({ placeholder: { options: { name: 'title', type: 'title', align: 'left', x: M, y: 0.7, w: 8.0, h: 0.75, margin: 0, valign: 'middle', fontFace: FONT, fontSize: 28, bold: true, color: BLACK }, text: 'Slide heading' } });
  logo(objects, { x: W - M - 0.42, y: 0.42, size: 0.42, onDark: false, withName: false });
  footer(objects, false);
}

// ---------------------------------------------------------------------------
// Layout 4: Header only
// ---------------------------------------------------------------------------
{
  const objects = [];
  contentHeader(objects);
  pres.defineSlideMaster({ title: 'CTS Header Only', background: { color: WHITE }, objects, slideNumber: slideNumber(false) });
}

// ---------------------------------------------------------------------------
// Layout 5: Header and text
// ---------------------------------------------------------------------------
{
  const objects = [];
  contentHeader(objects);
  objects.push({
    placeholder: {
      options: { name: 'body', type: 'body', x: M, y: 1.7, w: 8.9, h: 3.3, margin: 0, valign: 'top', fontFace: FONT, fontSize: 16, color: BLACK, paraSpaceAfter: 8, bullet: { indent: 18 } },
      text: 'Body text',
    },
  });
  pres.defineSlideMaster({ title: 'CTS Header and Text', background: { color: WHITE }, objects, slideNumber: slideNumber(false) });
}

// ---------------------------------------------------------------------------
// Example slides, one or two per layout
// ---------------------------------------------------------------------------
const LOGO_NOTE = findLogo(false)
  ? 'Logo from presentation/assets. The brand guidelines say the logo is always locked up with the promise "Seamless AV".'
  : 'The CTS mark on this layout is a placeholder. Drop the real logo into presentation/assets/cts-logo.png (and cts-logo-dark.png for dark slides) and rebuild, or replace it on the Slide Master (View > Slide Master).';

{
  const s = pres.addSlide({ masterName: 'CTS Cover (dark)' });
  s.addText('MONTHLY REPORTING', { placeholder: 'eyebrow' });
  s.addText('Financial Controller Pack', { placeholder: 'title' });
  s.addText('Month-end results, budget variances and the outlook for the quarter', { placeholder: 'subtitle' });
  s.addText('Presenter name, Finance', { placeholder: 'presenter' });
  s.addText('September 2026', { placeholder: 'date' });
  s.addNotes('Cover (dark) layout. ' + LOGO_NOTE);
}
{
  const s = pres.addSlide({ masterName: 'CTS Cover (light)' });
  s.addText('TRAINING', { placeholder: 'eyebrow' });
  s.addText('Payroll Processing Walkthrough', { placeholder: 'title' });
  s.addText('The fortnightly cycle from payroll notes to bank upload', { placeholder: 'subtitle' });
  s.addText('Presenter name, Finance', { placeholder: 'presenter' });
  s.addText('September 2026', { placeholder: 'date' });
  s.addNotes('Cover (light) layout. ' + LOGO_NOTE);
}
{
  const s = pres.addSlide({ masterName: 'CTS Section' });
  s.addText('01', { placeholder: 'number' });
  s.addText('Results for the month', { placeholder: 'title' });
  s.addText('Revenue, gross margin and overheads against budget', { placeholder: 'subtitle' });
  s.addNotes('Section divider layout. Change the number and title for each section.');
}
{
  const s = pres.addSlide({ masterName: 'CTS Section' });
  s.addText('02', { placeholder: 'number' });
  s.addText('Outlook and next steps', { placeholder: 'title' });
  s.addText('What changes next month and who owns each action', { placeholder: 'subtitle' });
}
{
  const s = pres.addSlide({ masterName: 'CTS Header Only' });
  s.addText('RESULTS FOR THE MONTH', { placeholder: 'section' });
  s.addText('Revenue by division, rolling twelve months', { placeholder: 'title' });
  // Example content so the layout is not empty: a native chart in the division colours.
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
  s.addNotes('Header only layout. Delete the sample chart and drop in your own chart, table or image. Chart values are illustrative only. Division colours: Production mint, Consulting lavender, Support rose.');
}
{
  const s = pres.addSlide({ masterName: 'CTS Header Only' });
  s.addText('SECTION NAME', { placeholder: 'section' });
  s.addText('Header only, empty body for your content', { placeholder: 'title' });
}
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
  s.addNotes('Header and text layout, bullet version. Keep to four or five bullets per slide.');
}
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
  s.addNotes('Header and text layout with a narrower text box on the left and a callout card on the right.');
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
