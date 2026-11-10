"""Bring the GL paste up to Xero.

Checked every account in the paste against Xero for July to September. Two
differ, both in August, and nothing else moves at all.

  Other Insurance        Xero 5,183.30  paste 5,073.18   paste 110.12 low
  Staff Entertainment    Xero 7,355.63  paste 7,620.26   paste 264.63 high

Net, the paste carries 154.51 of overhead that Xero no longer has, which is
why the pack reads the quarter as a 5,807.80 loss against Xero's 5,653.26.

This appends one line per account, dated 31 August 2026, taking the job,
cost centre and layout from an existing line on the same account so the
reporting tabs read them the same way. They are marked in the description, so
a clean re-paste of the Xero export replaces them without trace.

July and September are untouched, so the September P&L paste and the PL_Check
control are unaffected.
"""
import datetime
import sys
import openpyxl

EPOCH = datetime.date(1899, 12, 30)
WHEN = (datetime.date(2026, 8, 31) - EPOCH).days
NOTE = 'Brought to Xero as at 10 Nov 2026'

# account code -> (debit, credit)
FIX = {'62300': (110.12, 0.0),      # Other Insurance, up
       '65115': (0.0, 264.63)}      # Staff Entertainment, down


def main(path):
    wb = openpyxl.load_workbook(path)
    ws = wb['P&L Analysis']

    # a line on the same account, to copy the shape of
    template = {}
    for r in range(2, ws.max_row + 1):
        code = str(ws.cell(r, 1).value or '').strip()
        if code in FIX and code not in template:
            template[code] = r
        if len(template) == len(FIX):
            break

    missing = [c for c in FIX if c not in template]
    if missing:
        print('no existing line to copy for:', missing)
        return

    row = ws.max_row + 1
    for code, (dr, cr) in FIX.items():
        src = template[code]
        for c in range(1, 14):
            ws.cell(row, c)._style = ws.cell(src, c)._style
        ws.cell(row, 1).value = ws.cell(src, 1).value          # account code
        ws.cell(row, 2).value = ws.cell(src, 2).value          # account name
        ws.cell(row, 3).value = 'Manual Journal'
        ws.cell(row, 4).value = WHEN
        ws.cell(row, 5).value = None
        ws.cell(row, 6).value = dr
        ws.cell(row, 7).value = cr
        ws.cell(row, 8).value = ws.cell(src, 8).value          # job
        ws.cell(row, 9).value = None
        ws.cell(row, 10).value = None
        ws.cell(row, 11).value = NOTE
        ws.cell(row, 12).value = ws.cell(src, 12).value        # cost centres
        ws.cell(row, 13).value = ws.cell(src, 13).value
        print('  row %-5d %-7s %-24s debit %8.2f credit %8.2f'
              % (row, code, str(ws.cell(src, 2).value)[:24], dr, cr))
        row += 1

    wb.save(path)
    print('saved')


if __name__ == '__main__':
    main(sys.argv[1])
