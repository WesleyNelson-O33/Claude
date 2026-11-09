"""Correct the note on why the split differs.

The first version read as though the invoice had been split carelessly. It had
not. The invoice mirrors the quote line for line. The split was set when the
job was quoted, and it has not been recoded in Xero since.
"""
import openpyxl
from openpyxl.styles import Font

OUT = 'recon/CTS revenue against cost - Jul to Sep 2026.xlsx'
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')

WAS_NOW = {
 35: 'What happened. The invoice mirrors the quote line for line, so the '
     'invoice itself is not the error. The split was set when the job was '
     'quoted, with the value mostly on labour, while the supplier bills the '
     'cost mostly as equipment.',
 36: 'This has not been corrected in Xero. The coding below is still as it '
     'was invoiced.',
 37: 'What does not change. The invoice total, the cost and the job margin are '
     'the same in every version above. Only the split between the two lines '
     'moves.',
 38: 'What does change. On the company profit and loss, equipment reads as a '
     '23.5% loss for the quarter. This job is the whole of that.',
 39: 'What to decide. Whether to recode the invoice in Xero to sit against the '
     'cost, or to change how the job is quoted so the two agree from the '
     'start. The same question applies on every job quoted this way.',
}
JOB = {
 101: 'The yellow line is the one that reads wrong. 124,635.75 of equipment '
      'was bought and 61,303.17 was invoiced as equipment, because the invoice '
      'mirrors the quote and the quote put the value mostly on labour.',
 102: 'The reclassification moves the shortfall from labour to equipment. The '
      "job's invoiced total, its cost and its margin do not change. Only the "
      'split between the two lines does.',
 103: 'This puts equipment at nil margin, which is the least that has to move '
      'for the line to stop reading as a loss. On the job\'s own 20.2% the '
      'equipment line would be 156,107 and labour 49,245.',
 104: 'Not corrected in Xero. The lines above are still as invoiced.',
}

wb = openpyxl.load_workbook(OUT)
for tab, notes in (('Aware invoice, was and now', WAS_NOW), ('Aware Super job', JOB)):
    ws = wb[tab]
    for r, text in notes.items():
        cell = ws.cell(r, 1, text)
        cell.font = NOTE
    print(f'  {tab}: {len(notes)} notes rewritten')
wb.save(OUT)
print('saved', OUT)
