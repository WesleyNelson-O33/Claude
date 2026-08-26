/*
 * Builds a training manual .docx from a JSON spec.
 *
 *   node tools/build_manual.js manuals/specs/part2.json
 *
 * The spec carries the prose; this file owns the look, so every manual in
 * the series comes out consistent.
 */
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, ImageRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  Header, Footer, PageNumber, LevelFormat,
} = require('docx');

const ACCENT = '1F4E79';
const MUTED = '595959';
const CONTENT_DXA = 9360;

const specPath = process.argv[2];
if (!specPath) {
  console.error('usage: node tools/build_manual.js <spec.json>');
  process.exit(1);
}
const spec = JSON.parse(fs.readFileSync(specPath, 'utf8'));
const root = path.join(__dirname, '..');
const shotsDir = path.join(root, spec.screenshotDir);

function body(text) {
  return new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text, size: 22 })] });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: 'manual-bullets', level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, size: 22 })],
  });
}

function shot(file, caption) {
  const image = fs.readFileSync(path.join(shotsDir, file));
  // Frames are 1280 wide; scale to the printable width of an A4 page.
  const width = 600;
  const height = Math.round((722 / 1280) * width);
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 160, after: 60 },
      children: [new ImageRun({ type: 'png', data: image, transformation: { width, height } })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [new TextRun({ text: caption, italics: true, size: 18, color: MUTED })],
    }),
  ];
}

function callout(label, text) {
  return new Table({
    columnWidths: [CONTENT_DXA],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: 'D0D7E5' },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: 'D0D7E5' },
      left: { style: BorderStyle.SINGLE, size: 18, color: ACCENT },
      right: { style: BorderStyle.SINGLE, size: 2, color: 'D0D7E5' },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: CONTENT_DXA, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: 'F2F6FB' },
        margins: { top: 120, bottom: 120, left: 180, right: 180 },
        children: [new Paragraph({
          children: [
            new TextRun({ text: `${label}  `, bold: true, size: 22, color: ACCENT }),
            new TextRun({ text, size: 22 }),
          ],
        })],
      })],
    })],
  });
}

function keyValueTable(rows) {
  const left = 2900;
  const right = CONTENT_DXA - left;
  return new Table({
    columnWidths: [left, right],
    rows: rows.map(([key, value], index) => {
      const fill = index % 2 ? 'FFFFFF' : 'F7F9FC';
      const cell = (text, bold, width) => new TableCell({
        width: { size: width, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill },
        margins: { top: 80, bottom: 80, left: 140, right: 140 },
        children: [new Paragraph({ children: [new TextRun({ text, bold, size: 20 })] })],
      });
      return new TableRow({ children: [cell(key, true, left), cell(value, false, right)] });
    }),
  });
}

const spacer = () => new Paragraph({ spacing: { after: 200 }, children: [] });

const children = [
  new Paragraph({
    spacing: { before: 800, after: 60 },
    children: [new TextRun({ text: spec.eyebrow, bold: true, size: 24, color: MUTED, characterSpacing: 60 })],
  }),
  new Paragraph({
    spacing: { after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 8 } },
    children: [new TextRun({ text: spec.title, bold: true, size: 52, color: ACCENT })],
  }),
  new Paragraph({
    spacing: { before: 160, after: 480 },
    children: [new TextRun({ text: spec.subtitle, size: 22, color: MUTED })],
  }),
];

for (const block of spec.blocks) {
  switch (block.type) {
    case 'h1':
      children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 360, after: 160 },
        children: [new TextRun({ text: block.text, bold: true, size: 28, color: ACCENT })],
      }));
      break;
    case 'step':
      children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 360, after: 160 },
        children: [new TextRun({ text: `Step ${block.number}  ${block.title}`, bold: true, size: 28, color: ACCENT })],
      }));
      break;
    case 'h2':
      children.push(new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 120 },
        children: [new TextRun({ text: block.text, bold: true, size: 24, color: ACCENT })],
      }));
      break;
    case 'p': children.push(body(block.text)); break;
    case 'bullet': children.push(bullet(block.text)); break;
    case 'shot': children.push(...shot(block.file, block.caption)); break;
    case 'callout': children.push(callout(block.label, block.text)); children.push(spacer()); break;
    case 'table': children.push(keyValueTable(block.rows)); children.push(spacer()); break;
    case 'next':
      children.push(new Paragraph({
        spacing: { before: 400 },
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 12 } },
        children: [],
      }));
      children.push(new Paragraph({
        spacing: { before: 160 },
        children: [
          new TextRun({ text: 'Next: ', bold: true, size: 22, color: ACCENT }),
          new TextRun({ text: block.text, size: 22 }),
        ],
      }));
      break;
    default:
      throw new Error(`unknown block type: ${block.type}`);
  }
}

const doc = new Document({
  creator: 'Corporate Technology Services',
  title: spec.title,
  description: spec.subtitle,
  numbering: {
    config: [{
      reference: 'manual-bullets',
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: '•',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 460, hanging: 240 } } },
      }],
    }],
  },
  styles: { default: { document: { run: { font: 'Calibri', size: 22 } } } },
  sections: [{
    properties: { page: { margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } } },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: spec.runningHead, size: 16, color: MUTED })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: ['Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES], size: 16, color: MUTED })],
        })],
      }),
    },
    children,
  }],
});

const out = path.join(root, spec.output);
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log(`wrote ${out} (${(buf.length / 1024).toFixed(0)} KB, ${children.length} blocks)`);
});
