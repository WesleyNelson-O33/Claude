const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, AlignmentType } = require('docx');
const S = '/home/user/Claude/screenshots/';
const sizeOf = (p) => { const b = fs.readFileSync(p); // PNG header: width at 16, height at 20
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20), buf: b }; };
function img(name, label) {
  const { w, h, buf } = sizeOf(S + name);
  const maxW = 600; const scale = Math.min(1, maxW / w);
  return [
    new Paragraph({ spacing: { before: 60 }, children: [new TextRun({ text: label, italics: true, size: 18, color: '555555', font: 'Arial' })] }),
    new Paragraph({ spacing: { after: 160 }, children: [new ImageRun({ type: 'png', data: buf, transformation: { width: Math.round(w * scale), height: Math.round(h * scale) } })] }),
  ];
}
const P = (t, opts = {}) => new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: t, font: 'Arial', size: 22, ...opts })] });
const H = (t) => new Paragraph({ spacing: { before: 240, after: 80 }, children: [new TextRun({ text: t, font: 'Arial', size: 24, bold: true })] });
const children = [
  P('Subject: Bonus Leave - flowchart v2.0 vs 2026 policy document, and tracker basis', { bold: true }),
  P('Hi Graham,'),
  P('Thanks for sending the CTS Additional Leave Bonus Scheme flowchart v2.0 (04/09/25). As we agreed in our meeting, the only change from the 2017 scheme is that Bonus Leave is no longer paid out on termination, and v2.0 confirms that: every other step is the same as v1.0. The tracker applies the written 2026 policy where it goes further than the flowchart, as set out below. Service before the scheme started on 27/11/2017 counts towards the 3-year and 5-year milestones, but no day is earned for a milestone that fell before that date.'),
  ...img('00_v2_header.png', 'Flowchart v2.0 04/09/25'),
  P('The written Bonus Leave Scheme Policy & Procedure 2026 v1.0 (effective 08/09/2026) says more than the flowchart does, and in four places it would give a different result. I have listed them below with the two documents side by side and what the tracker does, so you can confirm the flowchart is the version we work to, or tell me the policy document overrides it.'),
  H('1. Who is eligible'),
  P('Flowchart v2.0: full-time employees only.'),
  ...img('01_2017_full_time_only.png', 'Flowchart v2.0'),
  P("2026 policy: part-time employees on 24 or more hours a week are now eligible, from 5 years' service."),
  ...img('02_2026_scope_part_time_added.png', '2026 policy, Scope (page 3-4)'),
  P("Tracker: follows the 2026 policy. Isaac Butterworth's part-time service (38 hours) from 17/09/2019 counts, giving him 9 days, the same as the HR tally. Tyrell Shortt (part-time, 38 hours, from 31/08/2021) earned his first day on 31/08/2026."),
  H('2. The 5-year anniversary'),
  P('Flowchart v2.0: one day on the anniversary, then one day every three months (two days only from the 6-year anniversary).'),
  ...img('03_2017_year5_one_day.png', 'Flowchart v2.0, 5 years or more'),
  P('2026 policy: two days on the 5-year anniversary itself.'),
  ...img('04_2026_year5_two_days.png', '2026 policy, Accrual Rules (page 5)'),
  P('Tracker: one day for a 5-year anniversary before 08/09/2026, two days from that date on. This matches the HR tally for everyone who has already passed 5 years.'),
  H('3. Part-time accrual rate (new in 2026)'),
  ...img('05_2026_part_time_accrual_new.png', '2026 policy, Eligible Part-Time Employees (page 5-6)'),
  P('Tracker: applied for part-time staff on 24 hours or more. Tyrell is the first to reach it.'),
  H('4. When service is counted from'),
  P('Flowchart v2.0: from the full-time employment start date, or the eligibility letter date if different.'),
  ...img('06_2017_service_from_FT_start_or_letter.png', 'Flowchart v2.0, 3 to 5 years'),
  P('2026 policy: from the original start date with CTS, paused during parental leave or leave without pay.'),
  ...img('07_2026_commencement_original_start.png', '2026 policy, Definitions (page 3)'),
  ...img('08_2026_parental_leave_pause_new.png', '2026 policy, Qualifying Service and Extended Leave of Absence (page 5)'),
  P('Tracker: service counts from the full-time start date (or the Eligibility Letter date if different), including service before 27/11/2017. Milestones that fell before the scheme start earn nothing; the first day is the first milestone on or after 27/11/2017.'),
  P("On this basis Jordan Sexty's EH balance matches the scheme exactly. Duncan Lugstein holds 16 hours more than the scheme gives because HR ran his accruals on a Feb/May/Aug/Nov cycle instead of his October anniversary. Daniel Sobkowski is 104 hours under: the scheme gives him 43 days from 27/11/2017 but EH only holds 30. I checked Daniel's and Jordan's personnel files and could not find an eligibility letter for either of them."),
  H('5. Termination (the agreed change, already in v2.0)'),
  P('Flowchart v2.0: all due entitlements including annual leave are paid out, Bonus Leave is not.'),
  ...img('09_2017_termination_paid_out.png', 'Flowchart v2.0, leaving CTS'),
  P('2026 policy: the same, all unused Bonus Leave is forfeited.'),
  ...img('10_2026_termination_forfeited.png', '2026 policy, Termination of Employment (page 8)'),
  P('Tracker: leavers are removed and their EH balance zeroed.'),
  P('Everything else is the same in both: first day at 3 years, one day every four months to year 5, five days a year after that, no pro-rata, and nothing before the scheme start on 27/11/2017.'),
  ...img('11_v2_title_rates.png', 'Flowchart v2.0, yearly rates'),
  P('Decisions I need from you: (1) confirm the 2026 policy document governs where it differs from the flowchart (part-time eligibility and two days at 5 years), which is how the tracker is set; (2) whether the 16 hours Duncan is over stands, and whether I post the 104 hours owed to Daniel. The eligibility letters I listed in my earlier email still stand.'),
  P('Thanks,'),
  P('Danica'),
];
const doc = new Document({ styles: { default: { document: { run: { font: 'Arial', size: 22 } } } },
  sections: [{ properties: { page: { margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } }, children }] });
Packer.toBuffer(doc).then(b => { fs.writeFileSync('/home/user/Claude/Bonus Leave - flowchart v2.0 vs 2026 policy - email to Graham.docx', b); console.log('written'); });
