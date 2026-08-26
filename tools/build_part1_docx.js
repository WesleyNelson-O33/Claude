const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, ImageRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  Header, Footer, PageNumber, LevelFormat, TableOfContents,
} = require('docx');

const SHOTS = path.join(__dirname, '..', 'manuals', 'screenshots', 'part1');
const OUT = path.join(__dirname, '..', 'manuals', 'Payroll Processing - Part 1 - Payroll Prep.docx');

// Extracted frames are 1280x722; scale to the printable width of an A4 page.
const IMG_W = 600;
const IMG_H = Math.round((722 / 1280) * IMG_W);

const ACCENT = '1F4E79';
const MUTED = '595959';

function shot(file, caption) {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 160, after: 60 },
      children: [new ImageRun({
        type: 'png',
        data: fs.readFileSync(path.join(SHOTS, file)),
        transformation: { width: IMG_W, height: IMG_H },
      })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
      children: [new TextRun({ text: caption, italics: true, size: 18, color: MUTED })],
    }),
  ];
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after === undefined ? 120 : opts.after },
    children: [new TextRun({ text, size: 22 })],
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: 'manual-bullets', level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, size: 22 })],
  });
}

function stepHeading(number, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    children: [new TextRun({ text: `Step ${number}  ${title}`, bold: true, size: 28, color: ACCENT })],
  });
}

function callout(label, text) {
  return new Table({
    columnWidths: [9360],
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
        width: { size: 9360, type: WidthType.DXA },
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

function pathTable(rows) {
  return new Table({
    columnWidths: [2600, 6760],
    rows: rows.map(([k, v], i) => new TableRow({
      children: [
        new TableCell({
          width: { size: 2600, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: i % 2 ? 'FFFFFF' : 'F7F9FC' },
          margins: { top: 80, bottom: 80, left: 140, right: 140 },
          children: [new Paragraph({ children: [new TextRun({ text: k, bold: true, size: 20 })] })],
        }),
        new TableCell({
          width: { size: 6760, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: i % 2 ? 'FFFFFF' : 'F7F9FC' },
          margins: { top: 80, bottom: 80, left: 140, right: 140 },
          children: [new Paragraph({ children: [new TextRun({ text: v, size: 20 })] })],
        }),
      ],
    })),
  });
}

const doc = new Document({
  creator: 'Corporate Technology Services',
  title: 'Payroll Processing - Part 1 - Payroll Prep',
  description: 'Training manual for fortnightly payroll preparation',
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
  styles: {
    default: { document: { run: { font: 'Calibri', size: 22 } } },
  },
  sections: [{
    properties: { page: { margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } } },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({
            text: 'Payroll Processing — Part 1: Payroll Prep',
            size: 16, color: MUTED,
          })],
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
    children: [
      new Paragraph({
        spacing: { before: 800, after: 60 },
        children: [new TextRun({ text: 'PAYROLL PROCESSING', bold: true, size: 24, color: MUTED, characterSpacing: 60 })],
      }),
      new Paragraph({
        spacing: { after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 8 } },
        children: [new TextRun({ text: 'Part 1 — Payroll Prep', bold: true, size: 52, color: ACCENT })],
      }),
      new Paragraph({
        spacing: { before: 160, after: 480 },
        children: [new TextRun({
          text: 'Corporate Technology Services  ·  Employment Hero Payroll  ·  Fortnightly cycle',
          size: 22, color: MUTED,
        })],
      }),

      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { after: 160 },
        children: [new TextRun({ text: 'About this guide', bold: true, size: 28, color: ACCENT })],
      }),
      body('Payroll at CTS runs on a fortnightly cycle. This guide covers Payroll Prep — everything that must be done before a pay run is created. It is the first of a six-part series.'),
      body('Work through it in order. Every step here must be complete before you move on to Part 2, because a pay run created from incomplete prep has to be scrapped and rebuilt.'),

      callout('Start early.', 'Begin payroll prep on the Friday before the Monday you process payroll. Prep is not a Monday-morning task — leaving it until the day means chasing approvals under time pressure.'),

      new Paragraph({ spacing: { after: 200 }, children: [] }),
      callout('Before you begin, make sure you have:', 'Access to the CTS Accounts mailbox, the Y drive (OneDrive), and Employment Hero payroll.'),

      stepHeading(1, 'Collect payroll notes across the fortnight'),
      body('Throughout the fortnight you will receive emails covering anything that changes someone’s pay. These include:'),
      bullet('Payroll queries'),
      bullet('Backpays — for example, hours worked at the wrong rate, or work not captured on a timesheet'),
      bullet('Additions and bonuses'),
      bullet('Leave approvals and bonus leave'),
      body('These arrive in one of two places: directly in your own inbox, or in the CTS Accounts inbox. Anything landing in the Accounts inbox is filed under the Payroll category.'),
      ...shot('02-payroll-category-inbox.png', 'Figure 1 — The CTS Accounts inbox grouped by category. Payroll items are collected under the Payroll category.'),
      callout('Collect as you go.', 'Do not leave this until prep day. Gathering a fortnight of scattered emails in one sitting is where items get missed.'),

      stepHeading(2, 'Collate the notes into a single email'),
      body('How you collate the notes is your choice, but the recommended method is to keep them in a single draft email as they arrive.'),
      body('The reason is the approval step: once prep is finished, the collated notes go to Duncan or Graham for approval. If they are already in an email, that approval is a single send with no rework.'),
      ...shot('01-payroll-notes-email.png', 'Figure 2 — A running draft of payroll notes. Each item records who is affected, what is owed, and the dates involved. Employee details are blurred in this guide.'),

      stepHeading(3, 'Copy the timesheet template into the fortnight folder'),
      body('The timesheet template lives on the Y drive. Navigate to it as follows:'),
      pathTable([
        ['Drive', 'Y drive — CTS NAS Accounts (Y:) Backup, via OneDrive'],
        ['Folder', '03 Payroll'],
        ['Sub-folder', '02 Payroll Related  ›  Payroll Timesheets'],
        ['Then', 'The relevant financial year'],
      ]),
      new Paragraph({ spacing: { after: 200 }, children: [] }),
      ...shot('03-y-drive-03-payroll.png', 'Figure 3 — The Y drive folder structure. Payroll sits under 03 Payroll.'),
      body('Inside the financial year folder you will find the timesheet template alongside a folder for each fortnight. Copy the template into the correct fortnight folder — do not edit the template itself.'),
      ...shot('04-timesheets-folder-template.png', 'Figure 4 — The financial year folder, containing the fortnight folders and the timesheet template.'),
      callout('Name it for the right fortnight.', 'In the example shown, the fortnight ending is 21 August, covering the period 8 August to 21 August.'),

      stepHeading(4, 'Populate the timesheet'),
      body('Transfer each payroll note into your copy of the template. There are two distinct places information goes, and the difference matters.'),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 120 },
        children: [new TextRun({ text: 'The Notes section', bold: true, size: 24, color: ACCENT })],
      }),
      body('Every payroll note goes in the Notes section against the relevant employee — regardless of whether it changes their pay. This is the record of what was raised and why.'),
      ...shot('06-timesheet-notes-column.png', 'Figure 5 — The Notes column, recording the detail behind each adjustment.'),

      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 200, after: 120 },
        children: [new TextRun({ text: 'Columns E to S', bold: true, size: 24, color: ACCENT })],
      }),
      body('Columns E through S carry the hours and values that actually drive the pay calculation — Normal Hours, In Lieu Taken, Public Holiday, Annual Leave Hours, and the overtime multipliers T1.25, T1.5 and T1.75.'),
      ...shot('05-timesheet-columns.png', 'Figure 6 — Columns E to S. Employee names and pay rates are blurred in this guide.'),
      callout('Only if it changes the pay.', 'Populate columns E to S only where the item affects the actual dollar value of the employee’s pay. A note that carries no pay impact belongs in the Notes section and nowhere else.'),

      stepHeading(5, 'Confirm all timesheets are approved in Employment Hero'),
      body('The last prep step is a gate: you cannot create a pay run while timesheets are still awaiting approval.'),
      body('In Employment Hero, go to Employee Management › Timesheet Approval, then set the filters:'),
      pathTable([
        ['Show timesheets for period', 'Fortnight Ending — set to the fortnight you are processing'],
        ['Status', 'Submitted'],
      ]),
      new Paragraph({ spacing: { after: 200 }, children: [] }),
      ...shot('07-approve-timesheets-pending.png', 'Figure 7 — Approve Timesheets, filtered to Submitted for the fortnight ending 21 August. Employee names are blurred in this guide.'),
      body('Read the result as follows:'),
      bullet('If timesheets are listed under Submitted, approvals are still outstanding. You cannot start the pay run — chase the approvals first.'),
      bullet('If the list is empty, everything has been approved and you are clear to proceed.'),
      callout('Missing Timesheets is a separate warning.', 'The banner lists employees who have not created a timesheet for the period at all. Follow these up as well — an approved-but-absent timesheet is still a gap in the pay run.'),

      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 400, after: 160 },
        children: [new TextRun({ text: 'Prep checklist', bold: true, size: 28, color: ACCENT })],
      }),
      body('Payroll prep is complete when all of the following are true:'),
      bullet('All payroll notes for the fortnight have been collected from both inboxes'),
      bullet('The notes are collated in a single email, ready to send to Duncan or Graham for approval'),
      bullet('The timesheet template has been copied into the correct fortnight folder'),
      bullet('Every note is recorded in the Notes section'),
      bullet('Columns E to S are populated for every item that changes the value of pay'),
      bullet('No timesheets remain under Submitted in Employment Hero'),
      bullet('Any missing timesheets have been followed up'),

      new Paragraph({
        spacing: { before: 400 },
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: ACCENT, space: 12 } },
        children: [],
      }),
      new Paragraph({
        spacing: { before: 160 },
        children: [
          new TextRun({ text: 'Next: ', bold: true, size: 22, color: ACCENT }),
          new TextRun({ text: 'Part 2 — Creating a New Pay Run.', size: 22 }),
        ],
      }),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log('wrote', OUT, `(${(buf.length / 1024).toFixed(0)} KB)`);
});
