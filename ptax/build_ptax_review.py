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

ITEMS = [
    # journal, item, tab, cell, now, should be, effect, narration
    ('JOURNAL 1  Payroll Tax NSW October 2026',
     'October timesheets are not loaded',
     'September 2026 Timesheets', 'the whole tab',
     "September units, 1,689 rows", 'October units', 'none on the total',
     "The split of the tax across jobs comes off this tab, and it still holds "
     "September. Both pivots read it, the NSW one and the other states one, so "
     "if I wrote the journal now every job line would be September's split "
     "with October's total on top. I need the October timesheet export before "
     "either journal can be written. Nothing else on this list is blocked by "
     "it, only the job lines."),
    (None,
     "September's 2,001.04 is still in the October wages formula",
     'Main & NSW', 'E3',
     '=169021.21-2001.04-E8', '=169021.21-E8', 99.35,
     "I left September's figure in the October formula. In September the "
     "2,001.04 was part of super, where row 5 reads 19,240.54 plus 2,001.04 "
     "and comes to the 21,241.58 that Employment Hero gives as total NSW "
     "super. In October I have subtracted the same 2,001.04 from wages "
     "instead, and I have not added it back anywhere, because row 5 is typed "
     "in as 20,086.18 with no formula. So NSW wages are 2,001.04 light and "
     "the tax is 99.35 light. Before I change it I need to look at the "
     "October Employment Hero report and see whether the 169,021.21 I typed "
     "already leaves the employer contributed super out. If it does the "
     "formula is right and only the look of it is wrong. If it does not, the "
     "2,001.04 comes out of the formula. Main & NSW cell E3."),
    (None,
     'No exempt per diems are deducted anywhere in October',
     'Main & NSW and the four state tabs', 'E3 on each',
     'nothing deducted', 'the October exempt per diems per state', None,
     "Per diems are exempt so I take them out of wages before the tax is "
     "worked out. I did that in July, 75 on New South Wales. I did it in "
     "August, 300 on Queensland and 150 each on South Australia, Victoria and "
     "Western Australia. I did it in September, 300 on New South Wales and "
     "525 on South Australia. In October I have taken none out, on any state. "
     "If October had per diems I am taxing them. I need the October "
     "Employment Hero report, the section headed payments in pay categories "
     "exempt, and then I deduct whatever it shows state by state the same way "
     "I did in September."),
    ('JOURNAL 2  Payroll Tax ACT/QLD/SA/VIC/WA October 2026 Accrual',
     'October timesheets, the same problem as above',
     'September 2026 Timesheets', 'the whole tab',
     'September units', 'October units', 'none on the total',
     "Same reasoning as the NSW journal. The twenty one job lines on this "
     "journal are a pivot of the September timesheets, so they are September's "
     "split. The 2,700 total is not affected because it is a fixed monthly "
     "accrual, only the split across the jobs is."),
    (None,
     'The accrual table still carries ACT at 300 a month',
     'Other States = Accrual', 'G3 and G8',
     'ACT 300, month total 3,000', 'ACT nil, month total 2,700', 'none, the journal is right',
     "I dropped the ACT accrual from August and my own note in cell C17 says "
     "so. The table at the top of the tab still has ACT at 300 every month, "
     "so it reads 3,000 a month and 36,000 for the year, when what I actually "
     "post is 2,700 a month. The journal is right and the table is wrong. I "
     "took ACT out of the table so it agrees with what I post."),
    (None,
     'One cent between the two allocation rows',
     'Other States = Accrual', 'F14 against F15',
     '413.49 on row 14, 413.50 on row 15', 'the same figure on both', 'one cent on a job line',
     "Row 14 says 413.49 for the DTTL Melbourne Tech FT 2 job and row 15 says "
     "413.50. The September journal used 413.49. One cent, but the two rows "
     "should not disagree, so I made them match."),
    ('NOT OCTOBER, BUT IT BREAKS THE YEAR TO DATE',
     'The June column formulas point at the state rows',
     'Main & NSW', 'M20, M21, M22',
     '=+M13+M19, =+M14+M20, =+M15+M21',
     'the same shape as the other months', 'none in October',
     "June has the wrong formulas. Row 20 should be the threshold, row 21 the "
     "taxable wages and row 22 the tax, and instead each one adds a state row "
     "to the row above it. They read nil now because June is empty, so it "
     "costs nothing yet, but it will give a wrong answer in June. I copied "
     "the formulas across from May."),
    (None,
     'The two annual totals do not cover the whole year',
     'Main & NSW', 'N20 and N23',
     '=SUM(B20:G20) and =SUM(B23:L23)', 'B to M on both', 'none in October',
     "The threshold total only adds July to December and the payable total "
     "only adds July to May. Both should run to June. I widened them."),
    (None,
     'A stray figure in the April tax',
     'Main & NSW', 'K23',
     '=+K22-1.1276', '=+K22', 'none in October',
     "April takes 1.1276 off the tax for no reason I can see. I do not know "
     "where it came from, so I am leaving it until I can find out rather than "
     "guess. If it was a rounding fix from last year it does not belong in "
     "FY2027."),
    (None,
     'Row 14 on the four state tabs is a sum, not a threshold',
     'QLD, SA, VIC and WA', 'E14 on each',
     '=SUM(E4:E13)', 'the threshold formula or nothing at all', 'none, nothing reads it',
     "The October cell adds rows 4 to 13, and row 13 is itself the total of "
     "rows 3 to 12, so it counts the super twice. On Queensland it reads "
     "16,041.50 against total wages of 14,489.10. Nothing reads row 14, so "
     "the tax is not affected and no journal changes. September's version is "
     "no better, it divides row 13 by itself and gives the full monthly "
     "threshold for every state. I would take row 14 off the state tabs "
     "altogether, because the threshold only applies to New South Wales."),
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
    band(js, 7, ['Line', 'October as the file stands', 'September, posted',
                 'Note', '', '', '', '', ''],
         [46, 20, 18, 52, 14, 14, 16, 70, 9], height=32)
    start = [
        ('NSW payroll tax, Main & NSW cell E23', as_is, 5902.75,
         'October is before the two corrections below'),
        ('Other states accrual, a fixed monthly figure', 2700.00, 2700.00,
         'ACT dropped from August, so 2,700 not 3,000'),
        ('Both journals together', as_is + 2700, 8602.75, ''),
    ]
    for i, (lbl, o, s, nt) in enumerate(start):
        r = 8 + i
        js.cell(r, 1, lbl).font = BODY
        js.cell(r, 2, o).number_format = MONEY
        js.cell(r, 3, s).number_format = MONEY
        js.cell(r, 4, nt).font = NOTE
        for k in (2, 3):
            js.cell(r, k).font = BOLD if i == 2 else BODY
        if i == 2:
            for k in range(1, 5):
                js.cell(r, k).fill, js.cell(r, k).border = TOTF, BOX

    para(js, 12, 'ENDING POSITION', 9, font=BOLD, height=18)
    para(js, 13, 'Still to do. The NSW tax becomes 5,897.48 once the 2,001.04 '
                 'is dealt with, and it moves again by whatever the October '
                 'exempt per diems come to. Neither journal can be written at '
                 'all until the October timesheets are loaded, because the job '
                 'lines are a pivot of that tab. The other states accrual '
                 'stays at 2,700.', 9, height=46)

    r = 15
    HDRS = ['Item', 'Tab', 'Cell to change', 'The file says now',
            'It should be', 'Effect on the tax', 'Tax after', 'What i did',
            'Done']
    for jnl, item, tab, cell, now, should, eff, narr in ITEMS:
        if jnl:
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
        if 'timesheet' in item.lower() or eff is None:
            for k in range(1, 10):
                js.cell(r, k).fill = RED
        elif isinstance(eff, float):
            for k in range(1, 10):
                js.cell(r, k).fill = AMBER
        else:
            for k in range(1, 10):
                js.cell(r, k).fill = BLUE
        js.row_dimensions[r].height = 118
        r += 1

    r += 1
    para(js, r, 'WHERE TO FIND EACH ITEM', 9, font=BOLD, height=18)
    r += 1
    band(js, r, ['Item', 'Tab', 'Cell', 'Colour', 'What the colour means',
                 '', '', '', ''], height=30)
    r += 1
    WHERE = [
        ('October timesheets', 'September 2026 Timesheets', 'the whole tab',
         'RED', 'Stops the journal being written at all'),
        ("September's 2,001.04", 'Main & NSW', 'E3', 'AMBER',
         'Changes the tax, and the amount is known'),
        ('Exempt per diems', 'Main & NSW, QLD, SA, VIC, WA', 'E3 on each',
         'RED', 'Needs the October Employment Hero report first'),
        ('ACT in the accrual table', 'Other States = Accrual', 'G3 and G8',
         'BLUE', 'Tidy up, the journal is already right'),
        ('The one cent', 'Other States = Accrual', 'F14 against F15', 'BLUE',
         'Tidy up, the journal is already right'),
        ('June formulas', 'Main & NSW', 'M20, M21, M22', 'BLUE',
         'Costs nothing yet, wrong from June'),
        ('The two annual totals', 'Main & NSW', 'N20 and N23', 'BLUE',
         'Year to date reads wrong'),
        ('The stray 1.1276', 'Main & NSW', 'K23', 'BLUE',
         'Left alone until its source is known'),
        ('Row 14 on the state tabs', 'QLD, SA, VIC, WA', 'E14 on each',
         'BLUE', 'Nothing reads it, no journal changes'),
    ]
    for i, row in enumerate(WHERE):
        n = r + i
        for k, v in enumerate(row, start=1):
            c = as_text(js, n, k, v)
            c.font = BODY
            c.border = BOX
            c.alignment = Alignment(wrap_text=True, vertical='top')
        fill = {'RED': RED, 'AMBER': AMBER, 'BLUE': BLUE}[row[3]]
        js.cell(n, 4).fill = fill
    r += len(WHERE) + 2

    para(js, r, 'LEGEND', 9, font=BOLD, height=18)
    for i, (f, t) in enumerate((
            (RED, 'RED. Stops the journal, or cannot be settled without the '
                  'October Employment Hero report.'),
            (AMBER, 'AMBER. Changes the tax, and the amount is worked out.'),
            (BLUE, 'BLUE. Wrong on the face of the sheet but no journal '
                   'changes.'))):
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
         'diems, plus total super. That gives the taxable wages for each '
         'state. All five states tie to my file to the cent. Then the New '
         'South Wales threshold and the 5.45 per cent rate give 5,902.75, '
         'which is the journal I posted. So the method is sound and anything '
         'that looks different in October is a mistake in October, not in the '
         'method.', 74),
        ('', '', 10),
        ('Then I found three things that matter.',
         'The timesheets in the file are September\'s. Both allocation pivots '
         'read that tab, so the job lines on both October journals would carry '
         'September\'s split. I cannot write either journal until the October '
         'timesheets are in.', 58),
        ('', 'I left September\'s 2,001.04 in the October wages formula. In '
             'September that figure was part of super. In October I took it off '
             'wages instead and did not put it back anywhere. That makes New '
             'South Wales wages 2,001.04 light and the tax 99.35 light. I need '
             'the October Employment Hero report to see whether the figure I '
             'typed already excludes it.', 74),
        ('', 'I deducted no exempt per diems in October, on any state. I '
             'deducted them in July, August and September. If October had per '
             'diems I am taxing them, and I will not know until I have the '
             'October report.', 58),
        ('', '', 10),
        ('Then six smaller things.',
         'The accrual table still shows ACT at 300 a month even though I '
         'dropped it in August, so the table says 3,000 a month against the '
         '2,700 I post. One cent between the two allocation rows on the other '
         'states tab. June\'s formulas on Main & NSW point at the state rows '
         'instead of the threshold. The two annual totals stop short of June. '
         'April takes a stray 1.1276 off the tax. And row 14 on the four state '
         'tabs adds a total to its own components, so it counts super twice.',
         88),
        ('', '', 10),
        ('What I am not doing.',
         'I am not guessing the October per diems and I am not guessing what '
         'the 2,001.04 should be. Both need the October Employment Hero '
         'report. I am also leaving the 1.1276 alone until I know where it '
         'came from. Nothing on this list is posted yet.', 58),
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
