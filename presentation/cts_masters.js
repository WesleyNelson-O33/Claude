/*
 * CTS slide masters, shared by every deck built in this folder.
 *
 *   const cts = require('./cts_masters');
 *   const pres = cts.newDeck('My deck');      // pptxgenjs instance with all 19 layouts defined
 *   const s = pres.addSlide({ masterName: 'CTS Header Only' });
 *   await cts.writeDeck(pres, '/path/out.pptx', { potx: false });
 *
 * Palette: the official CTS Seamless AV brand palette (brand-source/
 * cts_pantonevalues.pdf). Base charcoal plus three division colours, each
 * with a light and a dark variant. Greys and the mint tint are derived here
 * for footers and cards; they are not brand colours.
 *
 * Logos: presentation/assets (see assets/README.md for sources).
 */
const path = require('path');
const fs = require('fs');
const pptxgen = require('pptxgenjs');
const JSZip = require('jszip');

// Brand palette (CTS Seamless AV)
const C = {
  BLACK: '1A1A1A',     // Base charcoal (the site's --charcoal, the logo export and the strategy PDF pages)
  CHARCOAL2: '242424', // the site's --charcoal-2, for panels on charcoal
  MINT: '3FD0C9',      // Production (light), PMS 3255 C
  LAVENDER: '6D71FF',  // Consulting (light), PMS 2124 C
  ROSE: 'F33844',      // Support (light), PMS 1788 C
  TEAL: '005358',      // Production (dark), PMS 7476 C
  NAVY: '122B82',      // Consulting (dark), PMS 287 C
  BURGUNDY: 'A8052E',  // Support (dark), PMS 1945 C
  WHITE: 'FFFFFF',
  // Derived neutrals (not brand colours)
  GREY: '5F6670',      // captions, footers, subtitles on white
  SOFT: 'C9CED6',      // secondary text on charcoal or navy
  MINT_TINT: 'E8F9F8', // callout cards on white
  LINE: 'E3E6EA',      // chart gridlines, table rules
  LIGHT: 'F4F4F4',     // the site's --grey-100
};

const FONT = 'Calibri'; // Brand font is Tomato Grotesk; Calibri is the safe fallback that ships with Office.

// The three divisions: light colour for fills on dark, dark colour for text on white.
const DIVISIONS = [
  { key: 'production', name: 'Production', light: C.MINT, dark: C.TEAL, promise: 'Engaging Experiences' },
  { key: 'consulting', name: 'Consulting', light: C.LAVENDER, dark: C.NAVY, promise: 'End to End collaboration' },
  { key: 'support', name: 'Support', light: C.ROSE, dark: C.BURGUNDY, promise: 'Personalised Service' },
];

const W = 10;     // LAYOUT_16x9
const H = 5.625;
const M = 0.55;   // side margin

// ---------------------------------------------------------------------------
// Logo files
// ---------------------------------------------------------------------------
const ASSETS = path.join(__dirname, 'assets');
function logoFile(kind, onDark) {
  // kind: 'cts' | '20yrs' | 'production' | 'consulting' | 'support'
  // onDark: false = charcoal version for white; true = white version for charcoal;
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

// ---------------------------------------------------------------------------
// Shared pieces
// ---------------------------------------------------------------------------
function circle(objects, { x, y, d, color, transparency }) {
  objects.push({ rect: { x, y, w: d, h: d, fill: { color, transparency }, line: { color, width: 0, transparency: 100 } } });
}
// The three division colours overlapping: the cover motif.
function trio(objects, ox, oy, scale) {
  circle(objects, { x: ox, y: oy, d: 2.6 * scale, color: C.MINT, transparency: 15 });
  circle(objects, { x: ox + 1.35 * scale, y: oy + 1.0 * scale, d: 2.6 * scale, color: C.LAVENDER, transparency: 25 });
  circle(objects, { x: ox + 0.35 * scale, y: oy + 2.1 * scale, d: 2.6 * scale, color: C.ROSE, transparency: 30 });
}
function footer(objects, onDark, label) {
  objects.push({
    text: {
      text: label,
      options: { x: M, y: H - 0.42, w: 5, h: 0.25, margin: 0, valign: 'middle', fontFace: FONT, fontSize: 9, color: onDark ? C.SOFT : C.GREY },
    },
  });
}
function slideNumber(onDark) {
  return { x: W - M - 0.6, y: H - 0.42, w: 0.6, h: 0.25, fontFace: FONT, fontSize: 9, color: onDark ? C.SOFT : C.GREY, align: 'right' };
}
function ph(name, type, opts, text) {
  return { placeholder: { options: { name, type, margin: 0, fontFace: FONT, ...opts }, text } };
}

// ---------------------------------------------------------------------------
// Layout definitions
// ---------------------------------------------------------------------------
function coverText(objects, { textW, eyebrowColor, titleColor, subColor, metaColor, titleSize = 40, dateX = 6.5, dateW = 2.95 }) {
  objects.push(ph('eyebrow', 'body', { x: M, y: 2.0, w: textW, h: 0.35, fontSize: 12, bold: true, color: eyebrowColor, charSpacing: 3 }, 'PRESENTATION TYPE'));
  objects.push(ph('title', 'title', { align: 'left', x: M, y: 2.35, w: textW, h: 1.35, valign: 'top', fontSize: titleSize, bold: true, color: titleColor }, 'Presentation title'));
  objects.push(ph('subtitle', 'body', { x: M, y: 3.7, w: textW, h: 0.6, valign: 'top', fontSize: 18, color: subColor }, 'Subtitle or one-line description'));
  objects.push(ph('presenter', 'body', { x: M, y: 4.65, w: 4.0, h: 0.55, valign: 'bottom', fontSize: 12, color: metaColor }, 'Presenter name, role'));
  objects.push(ph('date', 'body', { x: dateX, y: 4.65, w: dateW, h: 0.55, valign: 'bottom', align: 'right', fontSize: 12, color: metaColor }, 'Month Year'));
}

function sectionBody(objects, { blockColor, numberColor, titleColor, subColor, labelColor, onDark, footerLabel }) {
  objects.push({ text: { text: 'SECTION', options: { x: M, y: 0.45, w: 3, h: 0.35, margin: 0, valign: 'middle', fontFace: FONT, fontSize: 11, bold: true, color: labelColor, charSpacing: 3 } } });
  objects.push({ rect: { x: M, y: 2.05, w: 1.4, h: 1.4, fill: { color: blockColor }, line: { color: blockColor, width: 0 } } });
  objects.push(ph('number', 'body', { x: M, y: 2.05, w: 1.4, h: 1.4, align: 'center', valign: 'middle', fontSize: 48, bold: true, color: numberColor }, '01'));
  objects.push(ph('title', 'title', { align: 'left', x: M + 1.75, y: 2.05, w: 6.6, h: 0.9, valign: 'bottom', fontSize: 36, bold: true, color: titleColor }, 'Section title'));
  objects.push(ph('subtitle', 'body', { x: M + 1.75, y: 2.98, w: 6.6, h: 0.5, valign: 'top', fontSize: 16, color: subColor }, 'One line describing what this section covers'));
  footer(objects, onDark, footerLabel);
}

function contentHeader(objects, { logoKind, labelColor, footerLabel }) {
  objects.push(ph('section', 'body', { x: M, y: 0.4, w: 7.0, h: 0.3, valign: 'middle', fontSize: 10, bold: true, color: labelColor, charSpacing: 3 }, 'SECTION NAME'));
  objects.push(ph('title', 'title', { align: 'left', x: M, y: 0.7, w: 7.6, h: 0.75, valign: 'middle', fontSize: 28, bold: true, color: C.BLACK }, 'Slide heading'));
  logo(objects, logoKind, { x: W - M, y: 0.42, h: 0.6, onDark: false, alignRight: true });
  footer(objects, false, footerLabel);
}
function bodyPlaceholder(objects) {
  objects.push(ph('body', 'body', { x: M, y: 1.7, w: 8.9, h: 3.3, valign: 'top', fontSize: 16, color: C.BLACK, paraSpaceAfter: 8, bullet: { indent: 18 } }, 'Body text'));
}

function defineMasters(pres) {
  const CTS_FOOTER = 'Corporate Technology Services  |  Seamless AV';

  // Covers -------------------------------------------------------------------
  {
    const objects = [];
    trio(objects, 6.9, -0.6, 1.45);
    logo(objects, 'cts', { x: M, y: 0.45, h: 0.85, onDark: true });
    coverText(objects, { textW: 6.2, eyebrowColor: C.MINT, titleColor: C.WHITE, subColor: C.SOFT, metaColor: C.SOFT });
    pres.defineSlideMaster({ title: 'CTS Cover (dark)', background: { color: C.BLACK }, objects });
  }
  {
    const objects = [];
    objects.push({ rect: { x: 6.9, y: 0, w: 3.1, h: H, fill: { color: C.BLACK }, line: { color: C.BLACK, width: 0 } } });
    trio(objects, 6.5, 0.4, 1.25);
    objects.push({ rect: { x: 0, y: 0, w: 6.9, h: H, fill: { color: C.WHITE }, line: { color: C.WHITE, width: 0 } } }); // mask
    logo(objects, 'cts', { x: M, y: 0.45, h: 0.85, onDark: false });
    coverText(objects, { textW: 5.9, eyebrowColor: C.NAVY, titleColor: C.BLACK, subColor: C.GREY, metaColor: C.GREY, titleSize: 36, dateX: 4.2, dateW: 2.4 });
    pres.defineSlideMaster({ title: 'CTS Cover (light)', background: { color: C.WHITE }, objects });
  }
  {
    const objects = [];
    circle(objects, { x: 6.6, y: 2.2, d: 5.2, color: C.NAVY, transparency: 55 });
    logo(objects, '20yrs', { x: M, y: 0.5, h: 1.05, onDark: true });
    coverText(objects, { textW: 6.4, eyebrowColor: C.MINT, titleColor: C.WHITE, subColor: C.SOFT, metaColor: C.SOFT });
    pres.defineSlideMaster({ title: 'CTS Cover (20 years)', background: { color: C.BLACK }, objects });
  }
  for (const d of DIVISIONS) {
    const objects = [];
    circle(objects, { x: 6.4, y: -1.6, d: 4.6, color: d.light, transparency: 10 });
    circle(objects, { x: 7.6, y: 2.4, d: 4.2, color: d.dark, transparency: 20 });
    logo(objects, d.key, { x: M, y: 0.45, h: 0.85, onDark: true });
    coverText(objects, { textW: 6.0, eyebrowColor: d.light, titleColor: C.WHITE, subColor: C.SOFT, metaColor: C.SOFT });
    pres.defineSlideMaster({ title: `CTS ${d.name} Cover`, background: { color: C.BLACK }, objects });
  }

  // Sections -----------------------------------------------------------------
  {
    const objects = [];
    circle(objects, { x: 7.0, y: 2.4, d: 4.6, color: C.LAVENDER, transparency: 55 });
    logo(objects, 'cts', { x: W - M, y: 0.42, h: 0.6, onDark: true, alignRight: true });
    sectionBody(objects, { blockColor: C.MINT, numberColor: C.BLACK, titleColor: C.WHITE, subColor: C.SOFT, labelColor: C.MINT, onDark: true, footerLabel: CTS_FOOTER });
    pres.defineSlideMaster({ title: 'CTS Section (navy)', background: { color: C.NAVY }, objects, slideNumber: slideNumber(true) });
  }
  {
    const objects = [];
    circle(objects, { x: 7.0, y: 2.4, d: 4.6, color: C.MINT, transparency: 80 });
    logo(objects, 'cts', { x: W - M, y: 0.42, h: 0.6, onDark: false, alignRight: true });
    sectionBody(objects, { blockColor: C.BLACK, numberColor: C.WHITE, titleColor: C.BLACK, subColor: C.GREY, labelColor: C.NAVY, onDark: false, footerLabel: CTS_FOOTER });
    pres.defineSlideMaster({ title: 'CTS Section (light)', background: { color: C.WHITE }, objects, slideNumber: slideNumber(false) });
  }
  for (const d of DIVISIONS) {
    const objects = [];
    circle(objects, { x: 7.0, y: 2.4, d: 4.6, color: d.light, transparency: 60 });
    logo(objects, d.key, { x: W - M, y: 0.42, h: 0.6, onDark: 'white', alignRight: true });
    sectionBody(objects, { blockColor: d.light, numberColor: C.BLACK, titleColor: C.WHITE, subColor: C.SOFT, labelColor: d.light, onDark: true, footerLabel: `CTS ${d.name}  |  ${d.promise}` });
    pres.defineSlideMaster({ title: `CTS ${d.name} Section`, background: { color: d.dark }, objects, slideNumber: slideNumber(true) });
  }

  // Content ------------------------------------------------------------------
  const CONTENT = [
    { key: 'cts', prefix: 'CTS', labelColor: C.NAVY, footerLabel: CTS_FOOTER },
    ...DIVISIONS.map((d) => ({ key: d.key, prefix: `CTS ${d.name}`, labelColor: d.dark, footerLabel: `CTS ${d.name}  |  ${d.promise}` })),
  ];
  for (const c of CONTENT) {
    {
      const objects = [];
      contentHeader(objects, { logoKind: c.key, labelColor: c.labelColor, footerLabel: c.footerLabel });
      pres.defineSlideMaster({ title: `${c.prefix} Header Only`, background: { color: C.WHITE }, objects, slideNumber: slideNumber(false) });
    }
    {
      const objects = [];
      contentHeader(objects, { logoKind: c.key, labelColor: c.labelColor, footerLabel: c.footerLabel });
      bodyPlaceholder(objects);
      pres.defineSlideMaster({ title: `${c.prefix} Header and Text`, background: { color: C.WHITE }, objects, slideNumber: slideNumber(false) });
    }
  }
}

// ---------------------------------------------------------------------------
// Deck lifecycle
// ---------------------------------------------------------------------------
function newDeck(title) {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_16x9'; // 10in x 5.625in
  pres.author = 'Corporate Technology Services';
  pres.company = 'Corporate Technology Services';
  pres.title = title;
  defineMasters(pres);
  return pres;
}

// Writes the deck, then post-processes the package:
//  - pptxgenjs draws every master `rect` as a rectangle; the decorative
//    circles are the only master shapes with a transparent fill, so turn
//    those into ellipses.
//  - optionally saves a .potx copy (same package, template content type).
async function writeDeck(pres, out, { potx = false } = {}) {
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
  if (!potx) return;
  const ct = await zip.file('[Content_Types].xml').async('string');
  zip.file('[Content_Types].xml', ct.replace(
    'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml'));
  const potxPath = out.replace(/\.pptx$/, '.potx');
  fs.writeFileSync(potxPath, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log('wrote', potxPath);
}

module.exports = { C, FONT, DIVISIONS, W, H, M, newDeck, writeDeck, defineMasters };
