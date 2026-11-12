"""Build the September 2026 CSIRO adjustment schedule, in the approved layout.

One adjustment: the Dicker Data cost on job 26082401 was never accrued against
INV-10648, which went out on 25 September. So September carries the revenue and
none of the cost.

Laid out the way the September WIP Adjustment Schedule was: Journal Edit
Schedule, What i did, then the raw tabs, then the checks. The raw tabs are
carried straight across from the October adjustment file, untouched.

    scripts/build_csiro_sep.py <october adjustment file> <out.xlsx>
"""
import copy
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

INK, AMBER, PANEL, BLUE = '242C39', 'FFF2CC', 'F4F4F4', 'DDEBF7'
PINK, GREEN = 'FFC7CE', 'C6EFCE'
MONEY = '#,##0.00;(#,##0.00)'
ROW_EQ, ROW_SV = 505, 506           # the two CSIRO rows on the schedule
EQ, SV = 1056.87, 312.96

NARR_EQ = (
    "CSIRO was invoiced in September, not October. INV-10648 went out on 25 September "
    "for 1,792.00 and was paid on 28 September, so the revenue is in September. The "
    "Dicker Data bill on PO-00016939 still has not arrived, so there is no cost "
    "anywhere on this job. The cost is earned in the month the client was invoiced, so "
    "it is accrued in September, which puts 51100 Equipment - Purchase on the debit "
    "side and 11300 on the credit side. I posted a new September journal for it and I "
    "entered 1,056.87 as a negative in the September column of the schedule. I took it "
    "out of the October column, because October has neither the revenue nor the cost. "
    "Sheet1 row 505.")
NARR_SV = "Same reasoning as CSIRO equipment. Sheet1 row 506."

JOURNAL = [
    ('51100 Equipment - Purchase',
     'CSIRO - HP Phone Mod - Unmodified Phone, Dicker Data PO-00016939', EQ, None),
    ('11300 Work in Progress',
     'CSIRO - HP Phone Mod - Unmodified Phone, Dicker Data PO-00016939', None, EQ),
    ('51400 Service - hires',
     'CSIRO - HP Phone Mod - Modification Process, Dicker Data PO-00016940', SV, None),
    ('11300 Work in Progress',
     'CSIRO - HP Phone Mod - Modification Process, Dicker Data PO-00016940', None, SV),
]

WHAT = """How I am correcting September and what I based it on
Status: the correction below is not yet done, in Xero or on the WIP schedule.

WHAT I FOUND
I had the CSIRO job wrong in the October adjustment. I wrote that CSIRO had not been
invoiced and took the two lines out of #72083 on that basis. That was only true of
October. The invoice went out in September.
INV-10648 is dated 25 September 2026, reference CSIRO - HP Poly CCX400 Phone
Modification. Net 1,792.00, GST 179.20, total 1,971.20. It was paid on 28 September.
Its first line reads Project #: 26082401, which is the job on the schedule.
It splits 1,410.00 to 42600 Equipment Sales and 382.00 to 42300.
The Dicker Data bill has still not arrived. The last Dicker Data bill in Xero is
PO-00016928 dated 31 August 2026. There is nothing for PO-00016939 or PO-00016940.
So there is 1,792.00 of revenue in September and no cost against it.

WHY SEPTEMBER AND NOT OCTOBER
This is the same basis I used in the October pack. Where a cost has been incurred and
the supplier has not billed us yet, and the client has been invoiced in the month, the
cost is earned in the month and is accrued, so the cost account sits on the debit side
and 11300 on the credit side.
The client was invoiced in September, so September is the month.
Taking the two lines out of the October journal was still right. October has neither
the revenue nor the cost.

WHAT MOVES
September cost of sales goes up by 1,369.83, 1,056.87 on 51100 and 312.96 on 51400.
11300 goes from a credit of 133,253.18 to a credit of 134,623.01.
September net result goes from a loss of 41,464 to a loss of 42,834.
Q1 goes from a loss of 5,653 to a loss of 7,023.

WHERE THE FIGURES CAME FROM
The invoice and the Dicker Data bills were read from Xero through the live connection
on 12 November 2026, not from a downloaded export. To check them: Business, then
Invoices, search INV-10648. And Contacts, Dicker Data, for the bills.
The raw tabs in this file are the ones from the October adjustment pack. The September
balances on them are still current, because every October correction sat in the October
column and none of them touched September.

WHAT IS NOT DONE
Nothing here is posted. The journal is written out on the Journal Edit Schedule ready
to type, and the two schedule rows still hold the old October figures."""

CHECKS = [
    ('Checks', ''),
    ('September revenue on job 26082401',
     'INV-10648, 25 September 2026, net 1,792.00. 1,410.00 to 42600 and 382.00 to 42300. '
     'Paid 28 September. Read from Xero on 12 November 2026.'),
    ('Dicker Data bill',
     'Not received. Last Dicker Data bill in Xero is PO-00016928 dated 31 August 2026. '
     'Nothing for PO-00016939 or PO-00016940.'),
    ('Cost anywhere else on the job',
     'None found. The job does not appear in the General Ledger Detail export used for '
     'the October pack.'),
    ('September 11300 agrees today',
     'The schedule and Xero both close September at a credit of 133,253.18, so this is '
     'not a variance. It is an accrual missing from both.'),
    ('', ''),
    ('Questions for you', ''),
    ('Is September closed',
     'If you would rather not reopen it, the same four lines go in at 31 October '
     'instead and the figures go in the October column as negatives. The cost then '
     'sits a month after the revenue.'),
    ('The deck',
     'Slides 2 and 3 carry Q1 figures read from the pack. Once this is posted and the '
     'pack is re-pasted they need rebuilding. That is one command, nothing is typed on '
     'them.'),
    ('The raw schedule tab',
     'It is the 30 October version, before your October corrections. The September '
     'balances on it are current. The October ones are not. Send me your current WIP '
     'file if you want the October column live too.'),
]


def head(ws, r, c, text, size=11, fill=INK, colour='FFFFFF'):
    cell = ws.cell(r, c, text)
    cell.font = Font(name='Arial', bold=True, size=size, color=colour)
    if fill:
        cell.fill = PatternFill('solid', fgColor=fill)
    return cell


def main():
    if len(sys.argv) < 3:
        print(__doc__.strip())
        return 1
    src, out = sys.argv[1], sys.argv[2]
    wb = openpyxl.load_workbook(src)
    for gone in ('Journal Edit Schedule', 'What i did ', 'What i did',
                 'Checks and Questions', 'Raw Xero GL Detail'):
        if gone in wb.sheetnames:
            del wb[gone]
    for need in ('Wip schedule', 'WIP Corrections', 'Raw Xero WIP 11300'):
        if need not in wb.sheetnames:
            print('STOPPED: the October file has no %r tab' % need)
            return 1

    # the corrections tab: move the two rows out of October and into September
    cor = wb['WIP Corrections']
    raw = wb['Wip schedule']
    # Column E is the typed source on these rows and CA is '=E505', so the
    # amount is moved by negating E, because the accrual credits WIP, and by
    # pointing the September column at E instead of the October one.
    for row, amt in ((ROW_EQ, EQ), (ROW_SV, SV)):
        if round(float(raw.cell(row, 5).value or 0), 2) != amt:
            print('STOPPED: Wip schedule E%d is %r, expected %s'
                  % (row, raw.cell(row, 5).value, amt))
            return 1
        if str(raw.cell(row, 79).value).strip() != '=E%d' % row:
            print('STOPPED: Wip schedule CA%d is %r, expected =E%d'
                  % (row, raw.cell(row, 79).value, row))
            return 1
        cor.cell(row, 5).value = -amt                       # E, the typed source
        cor.cell(row, 77).value = '=E%d' % row              # BY, September
        cor.cell(row, 78).value = '=+BX%d+BY%d' % (row, row)  # BZ, Sept balance
        cor.cell(row, 79).value = None                      # CA, October
        cor.cell(row, 80).value = '=+BZ%d' % row            # CB, Oct balance
        for c in (77, 79):
            cor.cell(row, c).fill = PatternFill('solid', fgColor=GREEN)
        raw.cell(row, 79).fill = PatternFill('solid', fgColor=PINK)
        for c in range(1, 6):
            raw.cell(row, c).fill = PatternFill('solid', fgColor=PINK)
            cor.cell(row, c).fill = PatternFill('solid', fgColor=GREEN)
    print('  schedule rows %d and %d moved from the October column to September'
          % (ROW_EQ, ROW_SV))

    ws = wb.create_sheet('Journal Edit Schedule', 0)
    ws.sheet_view.showGridLines = False
    for c, w in ((1, 3.0), (2, 40.0), (3, 11.0), (4, 36.0), (5, 13.0), (6, 13.0),
                 (7, 11.0), (8, 13.0), (9, 95.0), (10, 12.3)):
        ws.column_dimensions[L(c)].width = w

    head(ws, 1, 2, 'WIP 11300 — Journal Edit Schedule, September 2026, CSIRO HP Poly accrual', 12)
    ws.merge_cells('B1:F1')
    s = ws.cell(2, 2, 'Status: the correction below is not yet done, in Xero or on the WIP '
                      'schedule. The Done column is blank because nothing has been posted.')
    s.font = Font(name='Arial', italic=True, size=9, color=INK)
    s.fill = PatternFill('solid', fgColor=AMBER)

    head(ws, 3, 2, 'STARTING POSITION', 10)
    ws.cell(3, 4, 'Closing 30 Sep, from the raw tabs').font = Font(
        name='Arial', bold=True, size=9, color=INK)
    head(ws, 3, 7, 'ENDING POSITION, once the correction is done', 10)
    ws.merge_cells('B3:C3')
    ws.merge_cells('G3:J3')
    SEP = ("-'Raw Xero WIP 11300'!F8"
           "+SUMIFS('Raw Xero WIP 11300'!E9:E185,'Raw Xero WIP 11300'!A9:A185,"
           "\"<\"&DATE(2026,10,1))"
           "-SUMIFS('Raw Xero WIP 11300'!F9:F185,'Raw Xero WIP 11300'!A9:A185,"
           "\"<\"&DATE(2026,10,1))")
    rows = [
        (4, 'Raw Copy of WIP Schedule', "='Wip schedule'!BZ4",
         'Corrected WIP Sheet 30092026', "='WIP Corrections'!BZ4"),
        (5, 'Raw Xero WIP 11300', '=' + SEP,
         'Corrected Xero WIP 11300', '=E5-$H$23-$H$25'),
        (6, 'Variance to clear', '=ROUND(E4-E5,2)', 'Variance', '=ROUND(J4-J5,2)'),
    ]
    for r, lab, f, glab, gf in rows:
        ws.cell(r, 2, lab).font = Font(name='Arial', bold=(r == 6), size=9, color=INK)
        ws.cell(r, 5, f).number_format = MONEY
        ws.cell(r, 7, glab).font = Font(name='Arial', bold=(r == 6), size=9, color=INK)
        ws.cell(r, 10, gf).number_format = MONEY
    ws.cell(7, 7, 'Cost going into September').font = Font(name='Arial', bold=True, size=9, color=INK)
    ws.cell(7, 10, '=$H$23+$H$25').number_format = MONEY
    n = ws.cell(7, 2, 'Both openings are read off the raw tabs. September agrees today, so '
                      'the variance is nil. This is not a variance to clear, it is an '
                      'accrual missing from both the schedule and Xero.')
    n.font = Font(name='Arial', italic=True, size=8, color=INK)
    ws.merge_cells(start_row=7, start_column=2, end_row=7, end_column=6)
    src = ws.cell(9, 2,
                  'Raw tabs. Wip schedule is the Schedule Updated sheet of 2026-10_WIP.xlsx, '
                  'as it stood on 30 October before the October corrections. Raw Xero WIP 11300 '
                  'is the Work in Progress Transactions export for 1 September to 31 October, '
                  'pulled on 9 November 2026, also before the October corrections. September is '
                  'current on both, because every October correction sat in the October column '
                  'and none of them touched September. WIP Corrections carries the October '
                  'corrections you have already posted plus this September one, which is why it '
                  'closes October at 167,687.30 and the raw tab does not.')
    src.font = Font(name='Arial', italic=True, size=8, color=INK)
    src.alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=9, start_column=2, end_row=9, end_column=10)
    ws.row_dimensions[9].height = 46

    for c, h in ((2, 'Item'), (3, 'Journal'), (4, 'Line to change'), (5, 'Xero now'),
                 (6, 'Schedule says'), (7, 'Effect on\nvariance'), (8, 'Variance\nafter'),
                 (9, 'What to do in Xero'), (10, 'Done')):
        cell = head(ws, 8, c, h, 9)
        cell.alignment = Alignment(wrap_text=True, vertical='center')

    head(ws, 10, 2, 'NEW SEPTEMBER JOURNAL — CSIRO HP Poly, accrue the Dicker Data cost '
                    'against INV-10648', 10, PANEL, INK)
    ws.merge_cells('B10:J10')
    thin = Side(style='thin', color='B0B6BF')
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    for i, (r, item, line, row, narr) in enumerate((
            (11, 'CSIRO HP Poly phones equipment', '51100 Equipment - Purchase / 11300 WIP',
             ROW_EQ, NARR_EQ),
            (12, 'CSIRO HP Poly phones service', '51400 Service - hires / 11300 WIP',
             ROW_SV, NARR_SV))):
        ws.cell(r, 2, item)
        ws.cell(r, 3, 'NEW')
        ws.cell(r, 4, line)
        ws.cell(r, 5, 0).number_format = MONEY
        ws.cell(r, 6, "='Wip schedule'!CA%d" % row).number_format = MONEY
        # what the schedule's September moves by, less what Xero's moves by
        ws.cell(r, 7, "='WIP Corrections'!BY%d-'Wip schedule'!BY%d+$H$%d"
                % (row, row, 23 if row == ROW_EQ else 25)).number_format = MONEY
        ws.cell(r, 8, '=ROUND(E$6+SUM($G$11:$G%d),2)' % r).number_format = MONEY
        c9 = ws.cell(r, 9, narr)
        c9.alignment = Alignment(wrap_text=True, vertical='top')
        for c in range(2, 11):
            cell = ws.cell(r, c)
            cell.font = Font(name='Arial', size=9, color=INK)
            cell.border = box
            if c in (5, 6, 7, 8):
                cell.alignment = Alignment(horizontal='right', vertical='top')
            elif c != 9:
                cell.alignment = Alignment(vertical='top')
        ws.row_dimensions[r].height = 92 if r == 11 else 28

    head(ws, 15, 2, 'THE JOURNAL TO POST', 10)
    ws.merge_cells('B15:D15')
    for r, lab, val in ((16, 'Date', '30 September 2026'),
                        (17, 'Narration', 'CSIRO HP Poly phone modification, job 26082401. '
                                          'Accrue the Dicker Data cost against INV-10648 '
                                          'invoiced 25 September. Reverse when the bill is '
                                          'received.'),
                        (18, 'Job on every line', '26082401'),
                        (19, 'Cost centre on every line', 'INT'),
                        (20, 'Tax rate on every line', 'BAS Excluded')):
        ws.cell(r, 2, lab).font = Font(name='Arial', bold=True, size=9, color=INK)
        v = ws.cell(r, 4, val)
        v.font = Font(name='Arial', size=9, color=INK)
        v.alignment = Alignment(wrap_text=True, vertical='top')
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=9)
    for c, h in ((2, 'Account'), (4, 'Description'), (8, 'Debit'), (9, 'Credit')):
        head(ws, 22, c, h, 9)
    for i, (acct, desc, dr, cr) in enumerate(JOURNAL):
        r = 23 + i
        ws.cell(r, 2, acct).font = Font(name='Arial', size=9, color=INK)
        d = ws.cell(r, 4, desc)
        d.font = Font(name='Arial', size=9, color=INK)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
        for c, v in ((8, dr), (9, cr)):
            cell = ws.cell(r, c, v)
            cell.number_format = MONEY
            cell.font = Font(name='Arial', size=9, color=INK)
        for c in (2, 4, 8, 9):
            ws.cell(r, c).border = box
    r = 23 + len(JOURNAL)
    ws.cell(r, 2, 'Total').font = Font(name='Arial', bold=True, size=9, color=INK)
    for c in (8, 9):
        cell = ws.cell(r, c, '=SUM(%s23:%s%d)' % (L(c), L(c), r - 1))
        cell.number_format = MONEY
        cell.font = Font(name='Arial', bold=True, size=9, color=INK)
        cell.border = box

    base = r + 2
    head(ws, base, 2, 'WHERE TO FIND EACH ITEM', 10)
    ws.merge_cells(start_row=base, start_column=2, end_row=base, end_column=4)
    for c, h in ((2, 'Item'), (3, 'Sheet1 row'), (4, 'Xero tab row'), (5, 'Colour'),
                 (6, 'Meaning')):
        head(ws, base + 1, c, h, 9)
    finds = [
        ('=B11&" "&TEXT(%s,"$#,##0.00")' % EQ, ROW_EQ, 'not in Xero', 'BLUE',
         'A line is missing from both the schedule and Xero'),
        ('=B12&" "&TEXT(%s,"$#,##0.00")' % SV, ROW_SV, 'not in Xero', 'BLUE',
         'A line is missing from both the schedule and Xero'),
    ]
    for i, (item, srow, xrow, col, mean) in enumerate(finds):
        r = base + 2 + i
        ws.cell(r, 2, item)
        ws.cell(r, 3, srow).number_format = '0'
        ws.cell(r, 4, xrow)
        ws.cell(r, 5, col)
        ws.cell(r, 6, mean)
        for c in range(2, 7):
            ws.cell(r, c).font = Font(name='Arial', size=9, color=INK)
    for i, text in enumerate((
            'RED = change the journal in Xero.',
            'AMBER = Xero is right, change the schedule.',
            'BLUE = a line is missing from the schedule.',
            'On this one both colours would be blue. The accrual is missing from the '
            'schedule and from Xero, because it was never raised in September.')):
        cell = ws.cell(base + 5 + i, 2, text)
        cell.font = Font(name='Arial', italic=(i == 3), size=9, color=INK)

    wd = wb.create_sheet('What i did ', 1)
    wd.sheet_view.showGridLines = False
    wd.column_dimensions['A'].width = 108
    for i, line in enumerate(WHAT.split('\n'), 1):
        cell = wd.cell(i, 1, line)
        upper = line.isupper() and len(line) > 3
        cell.font = Font(name='Arial', bold=(i == 1 or upper), size=10 if i == 1 else 9,
                         color=INK)
        if i == 2:
            cell.fill = PatternFill('solid', fgColor=AMBER)

    ck = wb.create_sheet('Checks and Questions')
    ck.sheet_view.showGridLines = False
    ck.column_dimensions['A'].width = 38
    ck.column_dimensions['B'].width = 96
    for i, (a, b) in enumerate(CHECKS, 1):
        ca = ck.cell(i, 1, a)
        ca.font = Font(name='Arial', bold=True, size=9, color=INK)
        if a in ('Checks', 'Questions for you'):
            ca.fill = PatternFill('solid', fgColor=INK)
            ca.font = Font(name='Arial', bold=True, size=10, color='FFFFFF')
        cb = ck.cell(i, 2, b)
        cb.font = Font(name='Arial', size=9, color=INK)
        cb.alignment = Alignment(wrap_text=True, vertical='top')
        ck.row_dimensions[i].height = 30 if len(b) > 90 else 15

    wb.save(out)
    print('  written %s' % out)
    print('  tabs: %s' % ', '.join(wb.sheetnames))
    return 0


if __name__ == '__main__':
    sys.exit(main())
