/*
 * Binds every part into one master manual with a contents page.
 *   node tools/build_master.js manuals/specs/master.json
 *
 * The master spec lists the part specs in order; the parts themselves are
 * unchanged, so the bound document can never drift from the standalone ones.
 */
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, BorderStyle } = require('docx');
const lib = require('./manual_lib');

const specPath = process.argv[2];
if (!specPath) {
  console.error('usage: node tools/build_master.js <master-spec.json>');
  process.exit(1);
}

const root = path.join(__dirname, '..');
const master = JSON.parse(fs.readFileSync(specPath, 'utf8'));

const children = [
  new Paragraph({
    spacing: { before: 1200, after: 60 },
    children: [new TextRun({
      text: master.eyebrow, bold: true, size: 24,
      color: lib.MUTED, characterSpacing: 60,
    })],
  }),
  new Paragraph({
    spacing: { after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: lib.ACCENT, space: 8 } },
    children: [new TextRun({ text: master.title, bold: true, size: 56, color: lib.ACCENT })],
  }),
  new Paragraph({
    spacing: { before: 160, after: 400 },
    children: [new TextRun({ text: master.subtitle, size: 22, color: lib.MUTED })],
  }),
];

for (const line of master.intro || []) {
  children.push(new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text: line, size: 22 })],
  }));
}

children.push(...lib.contentsPage('Contents'));

// One shared counter so figures number continuously across the whole book.
const counter = { figure: 0 };
for (const partPath of master.parts) {
  const spec = JSON.parse(fs.readFileSync(path.join(root, partPath), 'utf8'));
  children.push(...lib.renderSpec(spec, { nested: true, root, counter }));
}

const doc = new Document({
  creator: 'Corporate Technology Services',
  title: master.title,
  description: master.subtitle,
  numbering: lib.BULLET_NUMBERING,
  styles: lib.DEFAULT_STYLES,
  sections: [lib.sectionShell(master.runningHead, children)],
});

const out = path.join(root, master.output);
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log(`wrote ${out} (${(buf.length / 1024).toFixed(0)} KB, ${counter.figure} figures)`);
});
