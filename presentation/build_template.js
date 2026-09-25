/*
 * Builds the CTS PowerPoint template.
 *   node presentation/build_template.js
 *
 * Writes presentation/CTS Presentation Template.pptx with five slide layouts
 * built into the Slide Master (two covers, a section divider, header only,
 * header and text) and one example slide per layout so the deck opens with
 * something to copy.
 *
 * Palette is the CTS house palette already used by the manuals and packs in
 * this repo (tools/manual_lib.js, reporting/build_pnl.py).
 */
const path = require('path');
const pptxgen = require('pptxgenjs');

const NAVY = '1F4E79';
const BLUE = '2E74B5';
const PALE = 'D9E2F3';
const GREY = '595959';
const LIGHT = 'F2F2F2';
const WELL = 'F7F9FC';
const INK = '262626';
const WHITE = 'FFFFFF';
const FONT = 'Calibri';

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

// Placeholder logo mark: navy rounded square with "CTS" plus the company name.
// Swap these for the real logo image on the Slide Master once you have the file.
function logo(objects, { x, y, size, dark, withName }) {
  objects.push({
    rect: {
      x, y, w: size, h: size, fill: { color: dark ? WHITE : NAVY },
      rectRadius: size * 0.2, shape: pres.ShapeType.roundRect, line: { color: dark ? WHITE : NAVY, width: 0 },
    },
  });
  objects.push({
    text: {
      text: 'CTS',
      options: {
        x, y, w: size, h: size, align: 'center', valign: 'middle', margin: 0,
        fontFace: FONT, fontSize: size * 26, bold: true, color: dark ? NAVY : WHITE, charSpacing: 1,
      },
    },
  });
  if (withName) {
    objects.push({
      text: {
        text: 'Corporate Technology Services',
        options: {
          x: x + size + 0.15, y, w: 4.2, h: size, valign: 'middle', margin: 0,
          fontFace: FONT, fontSize: 13, bold: true, color: dark ? WHITE : NAVY,
        },
      },
    });
  }
}

// Soft circles used as the cover motif.
function circle(objects, { x, y, d, color, transparency }) {
  objects.push({
    rect: { x, y, w: d, h: d, shape: pres.ShapeType.ellipse, fill: { color, transparency }, line: { color, width: 0, transparency: 100 } },
  });
}

function footer(objects, dark) {
  objects.push({
    text: {
      text: 'Corporate Technology Services',
      options: {
        x: M, y: H - 0.42, w: 4, h: 0.25, margin: 0, valign: 'middle',
        fontFace: FONT, fontSize: 9, color: dark ? PALE : GREY,
      },
    },
  });
}

const slideNumber = { x: W - M - 0.6, y: H - 0.42, w: 0.6, h: 0.25, fontFace: FONT, fontSize: 9, color: GREY, align: 'right' };

// ---------------------------------------------------------------------------
// Layout 1: Cover (dark)
// ---------------------------------------------------------------------------
{
  const objects = [];
  circle(objects, { x: 7.4, y: -1.5, d: 4.2, color: BLUE, transparency: 65 });
  circle(objects, { x: 8.1, y: 3.9, d: 3.4, color: BLUE, transparency: 78 });
  logo(objects, { x: M, y: 0.45, size: 0.5, dark: true, withName: true });
  objects.push({
    placeholder: {
      options: { name: 'eyebrow', type: 'body', x: M, y: 2.0, w: 7, h: 0.35, margin: 0, fontFace: FONT, fontSize: 12, bold: true, color: PALE, charSpacing: 3 },
      text: 'PRESENTATION TYPE',
    },
  });
  objects.push({
    placeholder: {
      options: { name: 'title', type: 'title', align: 'left', x: M, y: 2.35, w: 7.2, h: 1.35, margin: 0, valign: 'top', fontFace: FONT, fontSize: 40, bold: true, color: WHITE },
      text: 'Presentation title',
    },
  });
  objects.push({
    placeholder: {
      options: { name: 'subtitle', type: 'body', x: M, y: 3.7, w: 7.2, h: 0.6, margin: 0, valign: 'top', fontFace: FONT, fontSize: 18, color: PALE },
      text: 'Subtitle or one-line description',
    },
  });
  objects.push({
    placeholder: {
      options: { name: 'presenter', type: 'body', x: M, y: 4.65, w: 4.5, h: 0.55, margin: 0, valign: 'bottom', fontFace: FONT, fontSize: 12, color: PALE },
      text: 'Presenter name, role',
    },
  });
  objects.push({
    placeholder: {
      options: { name: 'date', type: 'body', x: 6.5, y: 4.65, w: 2.95, h: 0.55, margin: 0, valign: 'bottom', align: 'right', fontFace: FONT, fontSize: 12, color: PALE },
      text: 'Month Year',
    },
  });
  pres.defineSlideMaster({ title: 'CTS Cover (dark)', background: { color: NAVY }, objects });
}

// ---------------------------------------------------------------------------
// Layout 2: Cover (light)
// ---------------------------------------------------------------------------
{
  const objects = [];
  objects.push({ rect: { x: 6.9, y: 0, w: 3.1, h: H, fill: { color: NAVY }, line: { color: NAVY, width: 0 } } });
  circle(objects, { x: 5.9, y: 0.9, d: 3.0, color: BLUE, transparency: 55 });
  circle(objects, { x: 8.1, y: 3.4, d: 2.8, color: PALE, transparency: 75 });
  // white mask so the circles stay inside the navy panel on the left edge
  objects.push({ rect: { x: 0, y: 0, w: 6.9, h: H, fill: { color: WHITE }, line: { color: WHITE, width: 0 } } });
  logo(objects, { x: M, y: 0.45, size: 0.5, dark: false, withName: true });
  objects.push({
    placeholder: {
      options: { name: 'eyebrow', type: 'body', x: M, y: 2.0, w: 5.8, h: 0.35, margin: 0, fontFace: FONT, fontSize: 12, bold: true, color: BLUE, charSpacing: 3 },
      text: 'PRESENTATION TYPE',
    },
  });
  objects.push({
    placeholder: {
      options: { name: 'title', type: 'title', align: 'left', x: M, y: 2.35, w: 5.9, h: 1.35, margin: 0, valign: 'top', fontFace: FONT, fontSize: 36, bold: true, color: NAVY },
      text: 'Presentation title',
    },
  });
  objects.push({
    placeholder: {
      options: { name: 'subtitle', type: 'body', x: M, y: 3.7, w: 5.9, h: 0.6, margin: 0, valign: 'top', fontFace: FONT, fontSize: 16, color: GREY },
      text: 'Subtitle or one-line description',
    },
  });
  objects.push({
    placeholder: {
      options: { name: 'presenter', type: 'body', x: M, y: 4.65, w: 3.4, h: 0.55, margin: 0, valign: 'bottom', fontFace: FONT, fontSize: 12, color: GREY },
      text: 'Presenter name, role',
    },
  });
  objects.push({
    placeholder: {
      options: { name: 'date', type: 'body', x: 4.2, y: 4.65, w: 2.4, h: 0.55, margin: 0, valign: 'bottom', align: 'right', fontFace: FONT, fontSize: 12, color: GREY },
      text: 'Month Year',
    },
  });
  pres.defineSlideMaster({ title: 'CTS Cover (light)', background: { color: WHITE }, objects });
}

// ---------------------------------------------------------------------------
// Layout 3: Section divider
// ---------------------------------------------------------------------------
{
  const objects = [];
  circle(objects, { x: 7.3, y: 2.6, d: 4.4, color: WHITE, transparency: 40 });
  objects.push({
    text: {
      text: 'SECTION',
      options: { x: M, y: 0.45, w: 3, h: 0.35, margin: 0, valign: 'middle', fontFace: FONT, fontSize: 11, bold: true, color: BLUE, charSpacing: 3 },
    },
  });
  logo(objects, { x: W - M - 0.42, y: 0.42, size: 0.42, dark: false, withName: false });
  objects.push({
    rect: { x: M, y: 2.05, w: 1.4, h: 1.4, shape: pres.ShapeType.roundRect, rectRadius: 0.28, fill: { color: NAVY }, line: { color: NAVY, width: 0 } },
  });
  objects.push({
    placeholder: {
      options: { name: 'number', type: 'body', x: M, y: 2.05, w: 1.4, h: 1.4, margin: 0, align: 'center', valign: 'middle', fontFace: FONT, fontSize: 48, bold: true, color: WHITE },
      text: '01',
    },
  });
  objects.push({
    placeholder: {
      options: { name: 'title', type: 'title', align: 'left', x: M + 1.75, y: 2.05, w: 6.6, h: 0.9, margin: 0, valign: 'bottom', fontFace: FONT, fontSize: 36, bold: true, color: NAVY },
      text: 'Section title',
    },
  });
  objects.push({
    placeholder: {
      options: { name: 'subtitle', type: 'body', x: M + 1.75, y: 2.98, w: 6.6, h: 0.5, margin: 0, valign: 'top', fontFace: FONT, fontSize: 16, color: GREY },
      text: 'One line describing what this section covers',
    },
  });
  footer(objects, false);
  pres.defineSlideMaster({ title: 'CTS Section', background: { color: PALE }, objects, slideNumber });
}

// ---------------------------------------------------------------------------
// Shared header for the two content layouts
// ---------------------------------------------------------------------------
function contentHeader(objects) {
  objects.push({
    placeholder: {
      options: { name: 'section', type: 'body', x: M, y: 0.4, w: 7.5, h: 0.3, margin: 0, valign: 'middle', fontFace: FONT, fontSize: 10, bold: true, color: BLUE, charSpacing: 3 },
      text: 'SECTION NAME',
    },
  });
  objects.push({
    placeholder: {
      options: { name: 'title', type: 'title', align: 'left', x: M, y: 0.7, w: 8.0, h: 0.75, margin: 0, valign: 'middle', fontFace: FONT, fontSize: 28, bold: true, color: NAVY },
      text: 'Slide heading',
    },
  });
  logo(objects, { x: W - M - 0.42, y: 0.42, size: 0.42, dark: false, withName: false });
  footer(objects, false);
}

// ---------------------------------------------------------------------------
// Layout 4: Header only
// ---------------------------------------------------------------------------
{
  const objects = [];
  contentHeader(objects);
  pres.defineSlideMaster({ title: 'CTS Header Only', background: { color: WHITE }, objects, slideNumber });
}

// ---------------------------------------------------------------------------
// Layout 5: Header and text
// ---------------------------------------------------------------------------
{
  const objects = [];
  contentHeader(objects);
  objects.push({
    placeholder: {
      options: {
        name: 'body', type: 'body', x: M, y: 1.7, w: 8.9, h: 3.3, margin: 0, valign: 'top',
        fontFace: FONT, fontSize: 16, color: INK, paraSpaceAfter: 8, bullet: { indent: 18 },
      },
      text: 'Body text',
    },
  });
  pres.defineSlideMaster({ title: 'CTS Header and Text', background: { color: WHITE }, objects, slideNumber });
}

// ---------------------------------------------------------------------------
// Example slides, one or two per layout
// ---------------------------------------------------------------------------
const LOGO_NOTE = 'The CTS mark on this layout is a placeholder. Replace it with the real logo file on the Slide Master (View > Slide Master).';

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
  s.addText('Revenue by client, rolling twelve months', { placeholder: 'title' });
  // Example content so the layout is not empty: a native chart in the palette.
  s.addChart(pres.ChartType.bar, [
    { name: 'Revenue', labels: ['Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep'], values: [410, 425, 380, 360, 440, 470, 455, 490, 520, 505, 530, 560] },
  ], {
    x: M, y: 1.7, w: 8.9, h: 3.3, barDir: 'col', chartColors: [NAVY],
    showLegend: false, showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 9, dataLabelColor: GREY, dataLabelFontFace: FONT,
    catAxisLabelColor: GREY, catAxisLabelFontSize: 10, catAxisLabelFontFace: FONT, catGridLine: { style: 'none' },
    valAxisLabelColor: GREY, valAxisLabelFontSize: 10, valAxisLabelFontFace: FONT, valGridLine: { color: 'E1E6EF', size: 0.5 },
    valAxisLineShow: false, catAxisLineShow: false,
  });
  s.addNotes('Header only layout. Delete the sample chart and drop in your own chart, table or image. Chart values are illustrative only.');
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
  ], { x: M, y: 1.7, w: 4.9, h: 3.3, margin: 0, isTextBox: true, valign: 'top', fontFace: FONT, fontSize: 16, color: INK, paraSpaceAfter: 8, bullet: false });
  s.addShape(pres.ShapeType.roundRect, { x: 5.8, y: 1.7, w: 3.65, h: 3.3, rectRadius: 0.16, fill: { color: PALE }, line: { color: PALE, width: 0 } });
  s.addText('KEY POINT', { x: 6.1, y: 2.05, w: 3.05, h: 0.3, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 10, bold: true, color: BLUE, charSpacing: 2 });
  s.addText('A callout, figure or quote that supports the text on the left', { x: 6.1, y: 2.4, w: 3.05, h: 1.6, margin: 0, isTextBox: true, valign: 'top', fontFace: FONT, fontSize: 18, bold: true, color: NAVY });
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
const fs = require('fs');
const JSZip = require('jszip');

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
