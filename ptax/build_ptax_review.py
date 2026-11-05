"""October 2026 payroll tax, checked against September.

September is still in the file as column D, so it is the reference. Rebuilding
September from the Employment Hero report gives 5,902.75 to the cent, which is
the journal that was posted, so the method in the file is right and the
comparison below is against a method known to work:

    row 3  = EH Total Wages + Salary Sacrifice Super + Pre-Tax Deductions,
             less the exempt per diems for that state
    row 5  = EH Total Super
    row 12 = the two added, which equals the EH taxable line for that state
    NSW threshold = (NSW wages / all Australian wages) x 1.2m / 365 x days
    NSW tax = (all wages - threshold - interstate wages) x 5.45 per cent

All five states tie for September on that basis.

Laid out the way the September WIP Adjustment Schedule was approved, with two
changes, because nothing is in Xero yet. The columns say what the file holds
and what it should hold instead of what Xero holds, and there is no Xero
column. Everything else is the same: the narration column is Danica's own, the
Done column is left unticked, and the status line says the corrections are not
yet done.

Run with the October workbook, the September EH report, the September NSW
journal and the September other states journal, in that order.
"""
import csv
import datetime
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

OUT = 'ptax/CTS payroll tax October 2026 - review.xlsx'

NAVY, WHITE, GREY = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTF = PatternFill('solid', start_color=GREY, end_color=GREY)
RED = PatternFill('solid', start_color='FFFCE4E4', end_color='FFFCE4E4')
AMBER = PatternFill('solid', start_color='FFFFF3D6', end_color='FFFFF3D6')
BLUE = PatternFill('solid', start_color='FFE3EEF8', end_color='FFE3EEF8')
GREEN = PatternFill('solid', start_color='FFE4F2E4', end_color='FFE4F2E4')
RULE = Side(style='thin', color='FFC9CED6')
BOX = Border(top=RULE, bottom=RULE, left=RULE, right=RULE)
HEAD = Font(name='Calibri', sz=11, b=True, color=WHITE)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
BODY = Font(name='Calibri', sz=11, color=NAVY)
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')
TITLE = Font(name='Calibri', sz=14, b=True, color=NAVY)
WARN = Font(name='Calibri', sz=11, b=True, color='FFB00020')
MONEY = '#,##0.00;[RED](#,##0.00);"-"'
DATE = 'dd/mm/yyyy'

RATE, THRESH, DAYS_YR, DAYS_OCT = 0.0545, 1200000, 365, 31
INTER_OCT = 49139.64

# September, straight off the Employment Hero report for 1 to 30 September.
# Total Wages, Salary Sacrifice Super, Pre-Tax Deductions, Total Super, exempt
# per diems. The last column is what her file has on row 12 or 13.
EH_SEP = [
    ('NSW', 157739.98, 2725.88, 357.24, 21241.58, 300, 181764.68),
    ('QLD', 13175.12, 0.0, 0.0, 1581.01, 0, 14756.13),
    ('SA', 7036.87, 0.0, 0.0, 781.43, 525, 7293.30),
    ('VIC', 26418.15, 0.0, 0.0, 3068.43, 0, 29486.58),
    ('WA', 9596.29, 12.29, 0.0, 1143.58, 0, 10752.16),
]

# as lodged with Revenue NSW, version 2, 7 October 2026, against Main & NSW
# column D. The return is internally consistent on its own figures.
LODGED = [
    ('Salaries and wages', 158232.73, 160523.10),
    ('Employer superannuation', 24019.72, 21241.58),
    ('Total NSW wages', 182252.45, 181764.68),
    ('Interstate wages', 61800.00, 62288.17),
    ('Total Australian wages', 244052.45, 244052.85),
    ('Threshold entitlement', 73654.59, 73457.35),
    ('Tax payable', 5918.58, 5902.75),
]

ITEMS = [
    ('SEPTEMBER, ALREADY POSTED AND LODGED',
     'The return I lodged is 15.83 more than the journal I posted',
     'Main & NSW', 'D18 and D23',
     'Xero carries 5,902.75', 'Revenue NSW has 5,918.58', 15.83,
     "I lodged 5,918.58 with Revenue NSW on 7 October and I posted 5,902.75 "
     "to Xero, so Xero is 15.83 light on September. The cause is the "
     "interstate figure. The return used 61,800.00 and my file used "
     "62,288.17, and a lower interstate figure leaves more in the New South "
     "Wales net. The return also splits wages and super differently from my "
     "file, 158,232.73 and 24,019.72 against my 160,523.10 and 21,241.58, "
     "but that part nets out and does not change the tax. Two things to "
     "settle. Which interstate figure is right, and whether I top September "
     "up by 15.83 or let October carry it. The Lodged against the file tab "
     "sets the two side by side."),
    ('JOURNAL 1  Payroll Tax NSW October 2026',
     'October timesheets are not loaded',
     'September 2026 Timesheets', 'the whole tab',
     'September units, 1,689 rows', 'October units', 'none on the total',
     "The split of the tax across jobs comes off this tab and it still holds "
     "September. Both pivots read it, the New South Wales one and the other "
     "states one, so if I wrote the journal now every job line would be "
     "September's split with October's total on top. I need the October "
     "timesheet export before either journal can be written. Only the job "
     "lines are blocked by this, not the totals."),
    (None,
     "September's 2,001.04 is still in the October wages formula",
     'Main & NSW', 'E3',
     '=169021.21-2001.04-E8', '=169021.21-E8', 99.35,
     "I left September's figure in the October formula. In September the "
     "2,001.04 was part of super, where row 5 reads 19,240.54 plus 2,001.04 "
     "and comes to the 21,241.58 Employment Hero gives as total New South "
     "Wales super. In October I have taken the same 2,001.04 off wages "
     "instead and not added it back anywhere, because row 5 is typed in as "
     "20,086.18 with no formula. So New South Wales wages are 2,001.04 light "
     "and the tax is 99.35 light, which takes October from 5,798.13 to "
     "5,897.48. Before I change it I need the October Employment Hero report, "
     "to see whether the 169,021.21 I typed already leaves that super out. If "
     "it does, only the look of the formula is wrong."),
    (None,
     'No exempt per diems are deducted anywhere in October',
     'Main & NSW and the four state tabs', 'E3 on each',
     'nothing deducted', 'the October exempt per diems per state', None,
     "Per diems are exempt so I take them out of wages before the tax is "
     "worked out. I did that in July, 75 on New South Wales. In August, 300 "
     "on Queensland and 150 each on South Australia, Victoria and Western "
     "Australia. In September, 300 on New South Wales and 525 on South "
     "Australia. In October I have taken none out on any state. If October "
     "had per diems I am taxing them. I need the October Employment Hero "
     "report, the section headed payments in pay categories exempt, and then "
     "I deduct whatever it shows state by state."),
    ('JOURNAL 2  Payroll Tax ACT/QLD/SA/VIC/WA October 2026 Accrual',
     'October timesheets, the same problem as above',
     'September 2026 Timesheets', 'the whole tab',
     'September units', 'October units', 'none on the total',
     "Same reasoning as the New South Wales journal. The twenty one job lines "
     "are a pivot of the September timesheets. The 2,700 total is not "
     "affected because it is a fixed monthly accrual, only the split across "
     "the jobs is."),
    ('ALREADY CHANGED IN THE CORRECTED FILE I AM SENDING BACK',
     'One drag broke June. Row 19 was pulled down over rows 20, 21 and 22',
     'Main & NSW', 'M20, M21, M22',
     'all three were the Total Wages formula shifted down',
     "June's own threshold, taxable wages and tax", 'none in October',
     "This is one mistake, not three. The Total Wages formula on row 19 was "
     "dragged down the June column, and Excel recorded it as one shared "
     "formula covering E19 to M22. That is why June's threshold reads "
     "+M13+M19, its taxable wages +M14+M20 and its tax +M15+M21, all of them "
     "the row 19 formula shifted down a row at a time. They read nil now "
     "because June is empty, so it has cost nothing, but it would have given "
     "a wrong answer in June. I pulled the shared range back to row 19 and "
     "gave June its own three formulas."),
    (None,
     'The same drag on the state tabs. Row 13 was pulled down onto row 14',
     'QLD, SA, VIC, WA', 'E14 to L14 on each',
     '=SUM(E4:E13), which counts the super twice',
     'the formula July, August and September use', 'none, nothing reads it',
     "Same reasoning as June. The Total Wages formula on row 13 was dragged "
     "onto row 14, and the shared range covers both rows, so row 14 adds a "
     "total to its own parts. On Queensland that read 16,041.50 against total "
     "wages of 14,489.10. It runs from October all the way to May, not just "
     "October. Nothing reads row 14 so no journal changes and no tax changes. "
     "I put it back to the shape the first three months use. It would be "
     "better off the sheet altogether, because the 1.2 million threshold is "
     "New South Wales' and the other states are accrued at a fixed monthly "
     "figure, but taking a row out is a bigger change than I want to make "
     "without asking."),
    (None,
     'The year to date totals stopped short of June',
     'Main & NSW', 'N20 and N23',
     '=SUM(B20:G20) and =SUM(B23:L23)', 'B to M on both', 'none in October',
     "The threshold total only added July to December and the payable total "
     "only added July to May. I widened both to run to June."),
    (None,
     'Every empty month read #DIV/0! and broke the year to date',
     'Main & NSW and the state tabs', 'rows 20 to 23, and row 14',
     '32 cells reading #DIV/0!', 'nil for a month with no wages',
     'none in October',
     "A month with no wages divides by zero in the threshold, and that error "
     "ran along rows 20 to 23 and into every year to date total in column N, "
     "so the whole column was unusable. I wrapped the threshold in IFERROR so "
     "an empty month reads nil. The file went from 38 error cells to 6, and "
     "the 6 left are on the old Sheet1 tab, which I have not touched. The "
     "year to date payable now reads 23,165.57."),
    ('LEFT ALONE ON PURPOSE',
     'A stray figure in the April tax',
     'Main & NSW', 'K23',
     '=+K22-1.1276', '=+K22, if it should not be there', 'none in October',
     "April takes 1.1276 off the tax for no reason I can find, so I have left "
     "it rather than guess. It used to hide behind a #DIV/0!. Now that the "
     "empty months read nil it shows, and it is why the year to date payable "
     "reads 23,165.57 against 23,166.70 on the line above it. One cell fixes "
     "it if it does not belong in FY2027."),
    ('WHAT I TOLD YOU LAST TIME THAT WAS WRONG',
     'The one cent is deliberate, not a mistake',
     'Other States = Accrual', 'F14',
     '=ROUND(F15,2)-0.01', 'leave it exactly as it is', 'none',
     "I called this an error and it is not. Rounding each of the twenty one "
     "job lines to the cent pushes the allocation to 2,700.01, and taking a "
     "cent off the largest line brings it back to 2,700 so the variance check "
     "in E13 reads nil. That is the right way to do it. My mistake, not "
     "yours."),
    (None,
     'The ACT line in the accrual table is not wrong either',
     'Other States = Accrual', 'G3 and G8',
     'ACT 300, month total 3,000', 'leave it exactly as it is', 'none',
     "I called this an error too. The table at the top works the accrual out "
     "from the FY25 actual, and the posting rows below override ACT to nil "
     "with your own note in C17 saying the accrual was dropped from August "
     "2026. The journal posts 2,700 and that is right. The table is the "
     "starting point, not the answer. My mistake again."),
]



def as_text(ws, row, col, value):
    """Write a string as a string. A cell whose text begins with "=" would
    otherwise be saved as a formula, and the file opens showing #VALUE!."""
    c = ws.cell(row, col)
    c.value = value
    if isinstance(value, str) and value.startswith('='):
        c.data_type = 's'
    return c


def band(ws, row, labels, widths=None, height=42):
    for i, t in enumerate(labels, start=1):
        c = ws.cell(row, i, t)
        c.fill, c.font, c.border = BAND, HEAD, BOX
        c.alignment = Alignment(horizontal='center', wrap_text=True,
                                vertical='center')
    ws.row_dimensions[row].height = height
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[gl(i)].width = w


def para(ws, row, text, cols, font=None, height=30):
    c = ws.cell(row, 1, text)
    c.font = font or BODY
    c.alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    ws.row_dimensions[row].height = height


def nsw_tax(wages, inter=INTER_OCT, days=DAYS_OCT):
    total = wages + inter
    thr = (wages / total) * (THRESH / DAYS_YR * days)
    return (total - thr - inter) * RATE


def main():
    if len(sys.argv) != 5:
        print(__doc__)
        return 1
    poct, peh, pnsw, pother = sys.argv[1:5]

    as_is = nsw_tax(169021.21 - 2001.04 - 2000 + 20086.18 + 2000)
    fixed = nsw_tax(169021.21 - 2000 + 20086.18 + 2000)
    print(f'  October NSW as the file stands {as_is:,.2f}, '
          f'with the 2,001.04 back in {fixed:,.2f}')

    wb = openpyxl.Workbook()

    # ============================================ 1. Journal Edit Schedule
    js = wb.active
    js.title = 'Journal Edit Schedule'
    js['A1'] = 'October 2026 payroll tax, what to change before the journals go in'
    js['A1'].font = TITLE
    js['A2'] = 'Corporate Technology Services Pty Ltd'
    js['A2'].font = BODY
    para(js, 3, 'THE CORRECTIONS ON THIS SCHEDULE ARE NOT YET DONE. The Done '
                'column is blank on every line and the ending position below '
                'is still to do.', 9, font=WARN, height=32)
    para(js, 4, 'September is still in this file as column D, so it is the '
                'reference. September rebuilt from the Employment Hero report '
                'comes to 5,902.75, which is the journal that was posted, so '
                'the method in the file is right. The Method check tab shows '
                'that, state by state.', 9, font=NOTE, height=44)

    para(js, 6, 'STARTING POSITION', 9, font=BOLD, height=18)
    band(js, 7, ['Line', 'Amount', '', 'Note', '', '', '', '', ''],
         [50, 18, 12, 58, 46, 14, 14, 96, 9], height=32)
    start = [
        ('September NSW, lodged with Revenue NSW on 7 October', 5918.58,
         'version 2 of the return, worked out on actuals', True),
        ('September NSW, posted to Xero', 5902.75,
         'so Xero is 15.83 light on September', True),
        ('October NSW, as the file stands', as_is,
         'before the two corrections below', False),
        ('October other states accrual', 2700.00,
         'a fixed monthly figure, 3,000 less the ACT 300', False),
    ]
    for i, (lbl, amt, nt, hot) in enumerate(start):
        r = 8 + i
        js.cell(r, 1, lbl).font = BODY
        js.cell(r, 2, amt).number_format = MONEY
        js.cell(r, 2).font = BOLD
        js.cell(r, 4, nt).font = NOTE
        if hot:
            for k in range(1, 5):
                js.cell(r, k).fill = AMBER

    para(js, 12, 'ENDING POSITION', 9, font=BOLD, height=18)
    para(js, 13, 'Still to do. September needs a decision on the 15.83 '
                 'between the return and the journal. October NSW becomes '
                 '5,897.48 once the 2,001.04 is dealt with, and moves again '
                 'by whatever the October exempt per diems come to. Neither '
                 'October journal can be written until the October timesheets '
                 'are loaded, because the job lines are a pivot of that tab. '
                 'The other states accrual stays at 2,700. The formula faults '
                 'are already fixed in the corrected workbook.', 9, height=58)

    r = 15
    HDRS = ['Item', 'Tab', 'Cell to change', 'The file says now',
            'It should be', 'Effect on the tax', 'Tax after', 'What i did',
            'Done']
    done_now = withdrawn = False
    for jnl, item, tab, cell, now, should, eff, narr in ITEMS:
        if jnl:
            done_now = jnl.startswith('ALREADY CHANGED')
            withdrawn = jnl.startswith('WHAT I TOLD')
            para(js, r, jnl, 9, font=BOLD, height=20)
            js.cell(r, 1).fill = BLUE
            r += 1
            band(js, r, HDRS, [46, 20, 18, 52, 46, 14, 14, 96, 9])
            r += 1
        as_text(js, r, 1, item)
        as_text(js, r, 2, tab)
        as_text(js, r, 3, cell)
        as_text(js, r, 4, now)
        as_text(js, r, 5, should)
        if isinstance(eff, float):
            js.cell(r, 6, eff).number_format = MONEY
            js.cell(r, 7, as_is + eff).number_format = MONEY
        elif eff is None:
            as_text(js, r, 6, 'not worked out yet')
            as_text(js, r, 7, 'not worked out yet')
        else:
            as_text(js, r, 6, eff)
            as_text(js, r, 7, f'{as_is:,.2f}')
        as_text(js, r, 8, narr)
        js.cell(r, 9, '')
        for k in range(1, 10):
            c = js.cell(r, k)
            c.font = BODY
            c.border = BOX
            c.alignment = Alignment(wrap_text=True, vertical='top')
        js.cell(r, 9).fill = PatternFill('solid', start_color='FFFFFFFF',
                                         end_color='FFFFFFFF')
        if done_now:
            fill = GREEN
        elif withdrawn:
            fill = BLUE
        elif 'timesheet' in item.lower() or eff is None:
            fill = RED
        elif isinstance(eff, float):
            fill = AMBER
        else:
            fill = BLUE
        for k in range(1, 10):
            js.cell(r, k).fill = fill
        js.row_dimensions[r].height = 118
        r += 1

    r += 1
    para(js, r, 'WHERE TO FIND EACH ITEM', 9, font=BOLD, height=18)
    r += 1
    band(js, r, ['Item', 'Tab', 'Cell', 'Colour', 'What the colour means',
                 '', '', '', ''], height=30)
    r += 1
    WHERE = [
        ('The 15.83 on September', 'Main & NSW', 'D18 and D23', 'AMBER',
         'A month already posted and lodged, and the two disagree'),
        ('October timesheets', 'September 2026 Timesheets', 'the whole tab',
         'RED', 'Stops both October journals being written'),
        ("September's 2,001.04", 'Main & NSW', 'E3', 'AMBER',
         'Changes the October tax, and the amount is known'),
        ('Exempt per diems', 'Main & NSW, QLD, SA, VIC, WA', 'E3 on each',
         'RED', 'Needs the October Employment Hero report first'),
        ('The June drag', 'Main & NSW', 'M20, M21, M22', 'GREEN',
         'Already fixed in the corrected workbook'),
        ('The state tab drag', 'QLD, SA, VIC, WA', 'E14 to L14', 'GREEN',
         'Already fixed in the corrected workbook'),
        ('The year to date totals', 'Main & NSW', 'N20 and N23', 'GREEN',
         'Already fixed in the corrected workbook'),
        ('The #DIV/0! cells', 'Main & NSW and the state tabs',
         'rows 20 to 23, and row 14', 'GREEN',
         'Already fixed. 38 error cells down to 6'),
        ('The stray 1.1276', 'Main & NSW', 'K23', 'BLUE',
         'Left alone until its source is known'),
        ('The one cent', 'Other States = Accrual', 'F14', 'BLUE',
         'Not an error. I was wrong. Leave it'),
        ('ACT in the accrual table', 'Other States = Accrual', 'G3 and G8',
         'BLUE', 'Not an error. I was wrong. Leave it'),
    ]
    for i, row in enumerate(WHERE):
        n = r + i
        for k, v in enumerate(row, start=1):
            c = as_text(js, n, k, v)
            c.font = BODY
            c.border = BOX
            c.alignment = Alignment(wrap_text=True, vertical='top')
        fill = {'RED': RED, 'AMBER': AMBER, 'BLUE': BLUE,
                'GREEN': GREEN}[row[3]]
        js.cell(n, 4).fill = fill
    r += len(WHERE) + 2

    para(js, r, 'LEGEND', 9, font=BOLD, height=18)
    for i, (f, t) in enumerate((
            (RED, 'RED. Stops the journal, or cannot be settled without the '
                  'October Employment Hero report.'),
            (AMBER, 'AMBER. Changes the tax, and the amount is worked '
                    'out.'),
            (GREEN, 'GREEN. Already corrected in the workbook I am sending '
                    'back. Nothing for you to do.'),
            (BLUE, 'BLUE. Left as it is, on purpose or because I was wrong '
                   'to flag it.'))):
        n = r + 1 + i
        js.cell(n, 1).fill = f
        js.cell(n, 1).border = BOX
        js.cell(n, 2, t).font = BODY
        js.merge_cells(start_row=n, start_column=2, end_row=n, end_column=9)
    js.freeze_panes = 'A15'

    # ============================================ 2. What i did
    wd = wb.create_sheet('What i did')
    wd['A1'] = 'What i did'
    wd['A1'].font = TITLE
    lines = [
        ('', 'I am preparing the October 2026 payroll tax. September is still '
             'in the same file as column D, so I used it as the reference '
             'rather than working from memory.', 44),
        ('', '', 10),
        ('First I proved the method still works.',
         'I rebuilt September from the Employment Hero report. Wages plus '
         'salary sacrifice super plus pre tax deductions, less the exempt per '
         'diems, plus total super. All five states tie to my file to the '
         'cent, and the New South Wales threshold and the 5.45 per cent rate '
         'give 5,902.75, which is the journal I posted. So the method is '
         'sound and anything different in October is an October mistake.', 74),
        ('', '', 10),
        ('Then the lodged return turned up a problem in September.',
         'I lodged 5,918.58 with Revenue NSW on 7 October and I posted '
         '5,902.75 to Xero. Xero is 15.83 light. The cause is the interstate '
         'figure, 61,800.00 on the return against 62,288.17 in my file. '
         'September is already closed, so I need to decide whether to top it '
         'up or let October carry it, and I need to know which interstate '
         'figure is the right one.', 74),
        ('', '', 10),
        ('Then three things in October.',
         'The timesheets in the file are September\'s, and both allocation '
         'pivots read that tab, so the job lines on either October journal '
         'would be September\'s split. Neither journal can be written until '
         'the October timesheets are in.', 58),
        ('', 'I left September\'s 2,001.04 in the October wages formula. That '
             'makes New South Wales wages 2,001.04 light and the tax 99.35 '
             'light. I need the October Employment Hero report to see whether '
             'the figure I typed already excludes it.', 58),
        ('', 'I deducted no exempt per diems in October on any state, and I '
             'did deduct them in July, August and September. If October had '
             'per diems I am taxing them.', 44),
        ('', '', 10),
        ('Then I fixed the formulas, and it was one mistake twice.',
         'A total was dragged one row too far. On Main & NSW the Total Wages '
         'formula on row 19 was pulled down over rows 20, 21 and 22 in the '
         'June column, so June\'s threshold, taxable wages and tax were all '
         'the row 19 formula shifted down. On the four state tabs the Total '
         'Wages formula on row 13 was pulled down onto row 14 from October '
         'through to May, so row 14 counts the super twice. Neither one cost '
         'anything yet because the months are empty or nothing reads the row, '
         'but both were wrong. I also widened the two year to date totals to '
         'run to June, and wrapped the threshold in IFERROR so an empty month '
         'reads nil. That took the file from 38 error cells to 6, and the 6 '
         'left are on the old Sheet1 tab which I have not touched.', 102),
        ('', '', 10),
        ('Two things I called errors last time that are not.',
         'The one cent on the other states allocation is a deliberate '
         'rounding plug and it is right. The ACT line in the accrual table is '
         'the starting point, not the answer, and the posting rows below '
         'already override it to nil with my note saying so. Both were my '
         'mistake, not yours. Leave them alone.', 74),
        ('', '', 10),
        ('What I am not doing.',
         'I am not guessing the October per diems and I am not guessing what '
         'the 2,001.04 should be. Both need the October Employment Hero '
         'report. I am leaving the stray 1.1276 in April alone until I know '
         'where it came from. No October journal is posted.', 58),
    ]
    r = 3
    for head, text, h in lines:
        if head:
            js_c = wd.cell(r, 1, head)
            js_c.font = BOLD
            r += 1
        if text:
            c = wd.cell(r, 1, text)
            c.font = BODY
            c.alignment = Alignment(wrap_text=True, vertical='top')
            wd.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
            wd.row_dimensions[r].height = h
        r += 1
    for col, w in (('A', 24), ('B', 24), ('C', 24), ('D', 24), ('E', 24),
                   ('F', 24)):
        wd.column_dimensions[col].width = w

    # ============================================ 3. Method check
    mc = wb.create_sheet('Method check')
    mc['A1'] = 'September rebuilt from the Employment Hero report'
    mc['A1'].font = TITLE
    mc['A2'] = ('This is the proof that the method in the file is right, so '
                'the October comparison means something. Every figure in the '
                'first five columns is read off the September report.')
    mc['A2'].font = NOTE
    mc.merge_cells('A2:H2')
    mc.row_dimensions[2].height = 30
    band(mc, 4, ['State', 'EH total wages', 'EH salary sacrifice super',
                 'EH pre tax deductions', 'EH exempt per diems',
                 'EH total super', 'Taxable wages, rebuilt',
                 'Her file says', 'Ties'],
         [10, 16, 16, 16, 16, 16, 18, 16, 10])
    for i, (st, w, ss, pt, su, pd, hers) in enumerate(EH_SEP):
        r = 5 + i
        mc.cell(r, 1, st)
        for k, v in ((2, w), (3, ss), (4, pt), (5, pd), (6, su)):
            mc.cell(r, k, v).number_format = MONEY
        mc.cell(r, 7, f'=B{r}+C{r}+D{r}-E{r}+F{r}').number_format = MONEY
        mc.cell(r, 8, hers).number_format = MONEY
        mc.cell(r, 9, f'=IF(ROUND(G{r}-H{r},2)=0,"ties","CHECK")')
        for k in range(1, 10):
            mc.cell(r, k).font = BODY
    r = 5 + len(EH_SEP)
    mc.cell(r, 1, 'Total').font = BOLD
    for k in range(2, 9):
        c = mc.cell(r, k, f'=SUM({gl(k)}5:{gl(k)}{r-1})')
        c.number_format, c.font, c.fill, c.border = MONEY, BOLD, TOTF, BOX
    mc.cell(r, 9, f'=IF(ROUND(G{r}-H{r},2)=0,"ties","CHECK")').font = BOLD
    mc.cell(r, 1).fill, mc.cell(r, 1).border = TOTF, BOX

    r += 2
    para(mc, r, 'AND THEN THE NEW SOUTH WALES TAX', 9, font=BOLD, height=18)
    r += 1
    band(mc, r, ['Line', 'September', 'October as the file stands',
                 'October with the 2,001.04 back', '', '', '', '', ''],
         height=32)
    nsw_s, inter_s, days_s = 181764.68, 62288.17, 30
    rows = [
        ('NSW taxable wages', nsw_s, 187106.35, 189107.39),
        ('Interstate wages', inter_s, INTER_OCT, INTER_OCT),
        ('All Australian wages', nsw_s + inter_s, 236245.99, 238247.03),
        ('Days in the month', days_s, DAYS_OCT, DAYS_OCT),
    ]
    first = r + 1
    for i, (lbl, a, b, c) in enumerate(rows):
        n = first + i
        mc.cell(n, 1, lbl).font = BODY
        for k, v in ((2, a), (3, b), (4, c)):
            cell = mc.cell(n, k, v)
            cell.font = BODY
            cell.number_format = '#,##0' if 'Days' in lbl else MONEY
    n = first + 4
    mc.cell(n, 1, 'Threshold, apportioned').font = BODY
    for k in range(2, 5):
        c = mc.cell(n, k, f'=({gl(k)}{first}/{gl(k)}{first+2})'
                          f'*({THRESH}/{DAYS_YR}*{gl(k)}{first+3})')
        c.number_format, c.font = MONEY, BODY
    n += 1
    mc.cell(n, 1, 'Taxable after the threshold').font = BODY
    for k in range(2, 5):
        c = mc.cell(n, k, f'={gl(k)}{first+2}-{gl(k)}{first+4}'
                          f'-{gl(k)}{first+1}')
        c.number_format, c.font = MONEY, BODY
    n += 1
    mc.cell(n, 1, f'Payroll tax at {RATE*100:.2f} per cent').font = BOLD
    for k in range(2, 5):
        c = mc.cell(n, k, f'={gl(k)}{n-1}*{RATE}')
        c.number_format, c.font, c.fill, c.border = MONEY, BOLD, TOTF, BOX
    mc.cell(n, 1).fill, mc.cell(n, 1).border = TOTF, BOX
    n += 2
    mc.cell(n, 1, 'The September journal that was posted').font = BODY
    mc.cell(n, 2, 5902.75).number_format = MONEY
    mc.cell(n, 4, f'=IF(ROUND(B{n-2}-B{n},2)=0,"ties","CHECK")').font = BOLD
    n += 1
    mc.cell(n, 1, 'What the 2,001.04 is worth').font = BODY
    c = mc.cell(n, 4, f'=D{n-3}-C{n-3}')
    c.number_format, c.font = MONEY, BOLD

    # ============================================ 4. Lodged vs the file
    lg = wb.create_sheet('Lodged against the file', 3)
    lg['A1'] = 'September 2026, the lodged return against the working file'
    lg['A1'].font = TITLE
    lg['A2'] = ('Return history summary from Revenue NSW, lodgement version '
                '2, lodged 7 October 2026, calculated on actuals. Against '
                'Main & NSW column D.')
    lg['A2'].font = NOTE
    lg.merge_cells('A2:E2')
    lg.row_dimensions[2].height = 30
    band(lg, 4, ['Line', 'Lodged with Revenue NSW', 'The working file',
                 'Difference', 'What it means'],
         [30, 22, 20, 16, 64], height=32)
    for i, (label, lodged, filed) in enumerate(LODGED):
        r = 5 + i
        lg.cell(r, 1, label).font = BODY
        lg.cell(r, 2, lodged).number_format = MONEY
        lg.cell(r, 3, filed).number_format = MONEY
        lg.cell(r, 4, f'=B{r}-C{r}').number_format = MONEY
        for k in range(1, 5):
            lg.cell(r, k).font = BODY
        if label == 'Tax payable':
            for k in range(1, 6):
                lg.cell(r, k).fill = AMBER
                lg.cell(r, k).font = BOLD
                lg.cell(r, k).border = BOX
    notes = {
        'Salaries and wages': 'The return puts salary sacrifice super in the '
                              'super box, the file puts it in wages',
        'Employer superannuation': 'Same money, classified the other way. On '
                                   'its own this changes nothing',
        'Total NSW wages': 'The return has 487.77 more in New South Wales',
        'Interstate wages': 'And 488.17 less interstate. This is what moves '
                            'the tax',
        'Total Australian wages': 'Almost the same, so it is a split '
                                  'difference and not a missing payrun',
        'Threshold entitlement': 'Follows from the split, the return works '
                                 'its own threshold out correctly',
        'Tax payable': 'Revenue NSW has 5,918.58. Xero has 5,902.75. Xero is '
                       '15.83 light on September',
    }
    for i, (label, _l, _f) in enumerate(LODGED):
        c = lg.cell(5 + i, 5, notes.get(label, ''))
        c.font = BODY if label == 'Tax payable' else NOTE
        c.alignment = Alignment(wrap_text=True, vertical='top')
        lg.row_dimensions[5 + i].height = 30
    r = 5 + len(LODGED) + 1
    para(lg, r, 'THE CHECKS', 5, font=BOLD, height=18)
    for i, (t, f) in enumerate([
        ('The lodged return worked out again from its own figures',
         '=ROUND((B9-B10-B8)*0.0545,2)'),
        ('and that equals the tax it says is payable',
         '=IF(ROUND((B9-B10-B8)*0.0545-B11,2)=0,"yes","CHECK")'),
        ('The journal posted to Xero on 30 September', '=5902.75'),
        ('So Xero is light by', '=B11-5902.75'),
    ]):
        n = r + 1 + i
        lg.cell(n, 1, t).font = BODY
        c = lg.cell(n, 2, f)
        c.font = BOLD
        if 'yes' not in f:
            c.number_format = MONEY
    r = r + 6
    for i, t in enumerate([
        'TWO THINGS TO SETTLE',
        'Which interstate figure is right. The return used 61,800.00 and the '
        'file used 62,288.17. Whichever it is, the other one needs changing '
        'so next month starts from the same place.',
        'Whether September is topped up by 15.83 or October carries it. '
        'Either is fine, but the accrual account will sit 15.83 light until '
        'one of them happens.',
    ]):
        n = r + i
        para(lg, n, t, 5, font=BOLD if i == 0 else BODY,
             height=18 if i == 0 else 32)

    # ============================================ raw sources
    def add_csv(path, name, legend):
        rows = list(csv.reader(open(path, newline='', encoding='utf-8-sig')))
        ws = wb.create_sheet(name)
        for i, row in enumerate(rows, start=1):
            for j, v in enumerate(row, start=1):
                try:
                    v2 = float(v)
                except (TypeError, ValueError):
                    v2 = v
                c = ws.cell(i, j, v2)
                c.font = HEAD if i == 1 else BODY
                if i == 1:
                    c.fill, c.border = BAND, BOX
                if isinstance(v2, float):
                    c.number_format = MONEY
            if i > 1 and len(row) > 5:
                try:
                    amt = float(row[5])
                except (TypeError, ValueError):
                    amt = 0.0
                if amt < 0:
                    for j in range(1, len(row) + 1):
                        ws.cell(i, j).fill = AMBER
        for col, w in (('A', 54), ('B', 12), ('C', 52), ('D', 13),
                       ('E', 14), ('F', 12), ('G', 15), ('H', 40),
                       ('I', 14), ('J', 48)):
            ws.column_dimensions[col].width = w
        lg = len(rows) + 2
        ws.cell(lg, 1, 'LEGEND').font = BOLD
        ws.cell(lg + 1, 1).fill = AMBER
        ws.cell(lg + 1, 1).border = BOX
        ws.cell(lg + 1, 2, legend).font = BODY
        ws.freeze_panes = 'A2'

    add_csv(pnsw, 'Sep NSW journal',
            'The credit to 21440. The 99 debits to 65106 are the job split, '
            'and that split is what the October timesheets are needed for.')
    add_csv(pother, 'Sep other states journal',
            'The credit to 21440. 2,700, being 3,000 less the ACT 300 that '
            'was dropped in August.')

    ehv = openpyxl.load_workbook(peh, data_only=True)
    for t in ehv.sheetnames:
        src = ehv[t]
        ws = wb.create_sheet(f'Sep EH {t}'[:31])
        for row in src.iter_rows():
            for c in row:
                if c.value is None:
                    continue
                n = ws.cell(c.row, c.column, c.value)
                n.font = BODY
                if isinstance(c.value, float):
                    n.number_format = MONEY
        for r2 in range(1, src.max_row + 1):
            lbl = str(src.cell(r2, 1).value or '').lower()
            if 'per diem' in lbl or 'exempt' in lbl:
                for c2 in range(1, 8):
                    ws.cell(r2, c2).fill = RED
            if lbl in ('total wages', 'total super', 'grand total',
                       'total salary sacrifice super',
                       'total pre-tax deductions'):
                for c2 in range(1, 8):
                    ws.cell(r2, c2).fill = BLUE
        for col in 'ABCDEFG':
            ws.column_dimensions[col].width = 32 if col == 'A' else 15
        lg = src.max_row + 2
        ws.cell(lg, 1, 'LEGEND').font = BOLD
        for i, (f, t2) in enumerate((
                (RED, 'The exempt per diems. These come off wages before the '
                      'tax is worked out, and none were taken off in '
                      'October.'),
                (BLUE, 'The lines the file reads: wages, salary sacrifice, '
                       'pre tax deductions and super.'))):
            ws.cell(lg + 1 + i, 1).fill = f
            ws.cell(lg + 1 + i, 1).border = BOX
            ws.cell(lg + 1 + i, 2, t2).font = BODY

    wb.save(OUT)
    print(f'  written to {OUT}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
