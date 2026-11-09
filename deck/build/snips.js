const fs = require('fs');
const path = require('path');
const L = require('./lib.js');
const { C } = require('./theme.js');

const DIR = path.join(__dirname, '..', 'snips');

/** Width and height out of the PNG IHDR, so the files on disk are the truth. */
function pngSize(file) {
  const b = fs.readFileSync(file);
  return [b.readUInt32BE(16), b.readUInt32BE(20)];
}
const SIZES = Object.fromEntries(
  fs.readdirSync(DIR).filter((f) => f.endsWith('.png'))
    .map((f) => [f.replace(/\.png$/, ''), pngSize(path.join(DIR, f))]));

/**
 * Drop an Excel snip into a box, keeping its aspect ratio and sitting it at the
 * top of the box so a caption underneath always has somewhere to go.
 */
function placeSnip(slide, id, box) {
  const [pw, ph] = SIZES[id];
  if (!pw) throw new Error(`no snip rendered for ${id}`);
  const scale = Math.min(box.w / pw, box.h / ph);
  const w = pw * scale, h = ph * scale;
  slide.addImage({
    path: path.join(DIR, `${id}.png`),
    x: box.x + (box.w - w) / 2, y: box.y,
    w, h, objectName: `snip ${id}`,
  });
  return { w, h, bottom: box.y + h };
}

/** Same caption on every snip, so the reader always knows it is the pack. */
function source(slide, text, y, w = 6.10, x = 0.55) {
  slide.addText(text, {
    x, y, w, h: 0.26, isTextBox: true, margin: 0,
    fontSize: 8, italic: true, color: C.muted, valign: 'top',
  });
}

module.exports = { placeSnip, source, SIZES };
