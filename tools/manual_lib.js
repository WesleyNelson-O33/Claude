/*
 * Shared rendering for the payroll training manuals.
 *
 * The same spec renders two ways: standalone, where the part is the whole
 * document, and nested, where several parts are bound into one master and
 * every heading drops a level so the contents page reads as a hierarchy.
 */
const fs = require('fs');
const path = require('path');
const {
  Paragraph, TextRun, HeadingLevel, ImageRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  Header, Footer, PageNumber, LevelFormat, TableOfContents, PageBreak,
} = require('docx');

const ACCENT = '1F4E79';
const MUTED = '595959';
const CONTENT_DXA = 9360;

// Extracted frames are 1280 wide; scale to the printable width of an A4 page.
const IMG_WIDTH = 600;
const IMG_HEIGHT = Math.round((722 / 1280) * IMG_WIDTH);

const BULLET_NUMBERING = {
  config: [{
    reference: 'manual-bullets',
    levels: [{
      level: 0, format: LevelFormat.BULLET, text: '•',
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 460, hanging: 240 } } },
    }],
  }],
};

const DEFAULT_STYLES = { default: { document: { run: { font: 'Calibri', size: 22 } } } };

const PAGE_MARGIN = { top: 1080, bottom: 1080, left: 1080, right: 1080 };

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

function heading(text, level, size) {
  return new Paragraph({
    heading: level,
    spacing: { before: 360, after: 160 },
    children: [new TextRun({ text, bold: true, size, color: ACCENT })],
  });
}

function shot(dir, file, caption, figureNumber) {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 160, after: 60 },
      children: [new ImageRun({
        type: 'png',
        data: fs.readFileSync(path.join(dir, file)),
        transformation: { width: IMG_WIDTH, height: IMG_HEIGHT },
      })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [new TextRun({
        text: `Figure ${figureNumber} — ${caption}`,
        italics: true, size: 18, color: MUTED,
      })],
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

/** A question with the recording and words it came from, for citing in a meeting. */
function sourcedQuestions(items) {
  const out = [];
  for (const item of items) {
    out.push(new Paragraph({
      spacing: { before: 280, after: 60 },
      children: [
        new TextRun({ text: `${item.n}.  `, bold: true, size: 22, color: ACCENT }),
        new TextRun({ text: item.q, bold: true, size: 22 }),
      ],
    }));
    out.push(new Paragraph({
      spacing: { after: item.quote ? 60 : 120 },
      children: [new TextRun({ text: item.source, size: 18, color: MUTED, italics: true })],
    }));
    if (item.quote) {
      out.push(new Paragraph({
        indent: { left: 340 },
        spacing: { after: item.note ? 60 : 120 },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: 'D0D7E5', space: 10 } },
        children: [new TextRun({ text: `\u201c${item.quote}\u201d`, size: 20, italics: true })],
      }));
    }
    if (item.note) {
      out.push(new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun({ text: item.note, size: 20 })],
      }));
    }
  }
  return out;
}

/** Numbered questions with a blank cell beside each one to write the answer in. */
function questionTable(items, startNumber) {
  const numWidth = 620;
  const qWidth = 4400;
  const aWidth = CONTENT_DXA - numWidth - qWidth;
  const cell = (children, width, fill) => new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: { type: ShadingType.CLEAR, fill },
    margins: { top: 110, bottom: 110, left: 140, right: 140 },
    children,
  });
  const text = (t, bold, color) => new Paragraph({
    children: [new TextRun({ text: t, bold, size: 20, color })],
  });
  const header = new TableRow({
    tableHeader: true,
    children: [
      cell([text('#', true, 'FFFFFF')], numWidth, ACCENT),
      cell([text('Question', true, 'FFFFFF')], qWidth, ACCENT),
      cell([text('Answer', true, 'FFFFFF')], aWidth, ACCENT),
    ],
  });
  const rows = items.map((question, index) => {
    const fill = index % 2 ? 'FFFFFF' : 'F7F9FC';
    return new TableRow({
      children: [
        cell([text(String(startNumber + index), true)], numWidth, fill),
        cell([text(question, false)], qWidth, fill),
        cell([new Paragraph({ children: [] }), new Paragraph({ children: [] })], aWidth, fill),
      ],
    });
  });
  return new Table({ columnWidths: [numWidth, qWidth, aWidth], rows: [header, ...rows] });
}

function titleBlock(spec, nested) {
  if (nested) {
    // In the master the part title is the top-level heading, so it must carry
    // a real Heading style or it will not reach the contents page.
    return [
      new Paragraph({ children: [new PageBreak()] }),
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 240, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 8 } },
        children: [new TextRun({ text: spec.title, bold: true, size: 40, color: ACCENT })],
      }),
      new Paragraph({
        spacing: { before: 160, after: 320 },
        children: [new TextRun({ text: spec.subtitle, size: 20, color: MUTED })],
      }),
    ];
  }
  return [
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
}

/**
 * Render one spec's blocks.
 * @param {object} counter shared {figure: n} so numbering runs on across parts
 */
function renderSpec(spec, { nested = false, root = '.', counter } = {}) {
  const figures = counter || { figure: 0 };
  const shotsDir = path.join(root, spec.screenshotDir);
  const out = titleBlock(spec, nested);

  // Nested parts sit one level deeper than they do standalone.
  const majorLevel = nested ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_1;
  const minorLevel = nested ? HeadingLevel.HEADING_3 : HeadingLevel.HEADING_2;
  const majorSize = nested ? 26 : 28;
  const minorSize = nested ? 23 : 24;

  for (const block of spec.blocks) {
    switch (block.type) {
      case 'h1':
        out.push(heading(block.text, majorLevel, majorSize));
        break;
      case 'step':
        out.push(heading(`Step ${block.number}  ${block.title}`, majorLevel, majorSize));
        break;
      case 'h2':
        out.push(heading(block.text, minorLevel, minorSize));
        break;
      case 'p': out.push(body(block.text)); break;
      case 'bullet': out.push(bullet(block.text)); break;
      case 'shot':
        figures.figure += 1;
        out.push(...shot(shotsDir, block.file, block.caption, figures.figure));
        break;
      case 'callout':
        out.push(callout(block.label, block.text), spacer());
        break;
      case 'table':
        out.push(keyValueTable(block.rows), spacer());
        break;
      case 'questions':
        out.push(questionTable(block.items, block.start || 1), spacer());
        break;
      case 'sourced':
        out.push(...sourcedQuestions(block.items));
        break;
      case 'next':
        out.push(new Paragraph({
          spacing: { before: 400 },
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 12 } },
          children: [],
        }));
        out.push(new Paragraph({
          spacing: { before: 160 },
          children: [
            new TextRun({ text: nested ? 'Continues in: ' : 'Next: ', bold: true, size: 22, color: ACCENT }),
            new TextRun({ text: block.text, size: 22 }),
          ],
        }));
        break;
      default:
        throw new Error(`unknown block type: ${block.type}`);
    }
  }
  return out;
}

function contentsPage(title) {
  return [
    new Paragraph({
      spacing: { before: 240, after: 160 },
      children: [new TextRun({ text: title, bold: true, size: 32, color: ACCENT })],
    }),
    new Paragraph({
      spacing: { after: 200 },
      children: [new TextRun({
        text: 'If this page looks empty, Word has not yet built it. Right-click anywhere on it and choose Update Field, or answer Yes when Word offers to update fields on opening.',
        italics: true, size: 18, color: MUTED,
      })],
    }),
    new TableOfContents('Contents', { hyperlink: true, headingStyleRange: '1-3' }),
  ];
}

function sectionShell(runningHead, children) {
  return {
    properties: { page: { margin: PAGE_MARGIN } },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: runningHead, size: 16, color: MUTED })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            children: ['Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES],
            size: 16, color: MUTED,
          })],
        })],
      }),
    },
    children,
  };
}

module.exports = {
  renderSpec, contentsPage, sectionShell,
  BULLET_NUMBERING, DEFAULT_STYLES, ACCENT, MUTED,
};
