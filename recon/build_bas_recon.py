"""BAS lodged against Xero, January 2026 to September 2026.

Three quarters. The BAS G1 figures are typed from the lodged statements as
supplied. The Xero figures are typed from the Xero profit and loss, pulled
17 November 2026, on both cash and accrual basis, trading income only.
Everything else on the sheet is a formula pointing at those two blocks.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

OUT = 'recon/CTS BAS to Xero - Jan to Sep 2026.xlsx'

NAVY, WHITE, GREY, TEAL = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4', 'FF005358'
BLUE = 'FF0070C0'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTF = PatternFill('solid', start_color=GREY, end_color=GREY)
RULE = Side(style='thin', color='FFC9CED6')
HEAD = Font(name='Calibri', sz=11, b=True, color=WHITE)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
BODY = Font(name='Calibri', sz=11, color=NAVY)
INPUT = Font(name='Calibri', sz=11, color=BLUE)
TITLE = Font(name='Calibri', sz=16, b=True, color=NAVY)
SUB = Font(name='Calibri', sz=11, color='FF6B7585')
SECT = Font(name='Calibri', sz=12, b=True, color=TEAL)
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')
MONEY = '#,##0.00;[RED](#,##0.00);"-"'
PCT = '0.00%;[RED](0.00%);"-"'

# quarter, BAS G1 incl GST, Xero cash ex GST, Xero accrual ex GST, other income
QTRS = [
    ('Jan 2026 to Mar 2026', 1324425.00, 1210640.43, 1303788.88, 1514.36),
    ('Apr 2026 to Jun 2026', 1678257.00, 1526047.25, 1685817.62, 9788.27),
    ('Jul 2026 to Sep 2026', 1909366.00, 1735815.50, 1624459.78, 3400.20),
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'BAS to Xero'
ws.sheet_view.showGridLines = False


def put(r, c, v, font=BODY, fmt=None, fill=None, align=None, border=False):
    cell = ws.cell(r, c)
    cell.value = v
    cell.font = font
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = Alignment(horizontal=align, vertical='center',
                                   wrap_text=(align == 'centre'))
    if border:
        cell.border = Border(top=RULE, bottom=RULE, left=RULE, right=RULE)
    return cell


def bandrow(r, cols, labels):
    for i, lab in enumerate(labels):
        cell = put(r, cols + i, lab, HEAD, fill=BAND,
                   align='left' if i == 0 else 'right')
        cell.alignment = Alignment(horizontal='left' if i == 0 else 'right',
                                   vertical='center', wrap_text=True)
    ws.row_dimensions[r].height = 42


put(1, 1, 'BAS lodged against Xero', TITLE)
put(2, 1, 'Three quarters, January 2026 to September 2026. '
          'Corporate Technology Services Pty Ltd.', SUB)

put(4, 1, 'Where the figures came from', SECT)
put(5, 1, 'BAS G1 is read off the lodged activity statements attached at the bottom '
          'of this sheet. Both say GST accounting method Cash and G1 includes GST Yes. '
          'Typed once, it does not recalculate.', NOTE)
put(5, 2, '', NOTE)
put(6, 1, 'Xero is typed from the Xero profit and loss, pulled 17 November 2026. '
          'Trading income only, cash basis and accrual basis, both excluding GST. '
          'Typed once, it does not recalculate.', NOTE)
put(7, 1, 'Every other figure on this sheet is a formula pointing at those two blocks. '
          'Blue figures are typed. Navy figures are calculated.', NOTE)

put(9, 1, 'GST rate', BOLD)
put(9, 2, 0.10, INPUT, '0%').alignment = Alignment(horizontal='left')
put(9, 3, 'Input. Change it here and the whole sheet follows.', NOTE)

# ---------------------------------------------------------------- main block
put(11, 1, 'Cash basis, the basis the BAS is lodged on', SECT)
bandrow(12, 1, ['Quarter', 'BAS G1\nincl GST', 'BAS G1\nex GST',
                'Xero cash\nex GST', 'Variance\nex GST', 'Variance %',
                'Xero cash\nincl GST', 'Variance\nincl GST'])

first, last = 13, 12 + len(QTRS)
for i, (name, bas, cash, accr, oth) in enumerate(QTRS):
    r = first + i
    put(r, 1, name, BODY, border=True)
    put(r, 2, bas, INPUT, MONEY, border=True)
    put(r, 3, '=B%d/(1+$B$9)' % r, BODY, MONEY, border=True)
    put(r, 4, cash, INPUT, MONEY, border=True)
    put(r, 5, '=D%d-C%d' % (r, r), BODY, MONEY, border=True)
    put(r, 6, '=IF(D%d=0,"",E%d/D%d)' % (r, r, r), BODY, PCT, border=True)
    put(r, 7, '=D%d*(1+$B$9)' % r, BODY, MONEY, border=True)
    put(r, 8, '=B%d-G%d' % (r, r), BODY, MONEY, border=True)

t = last + 1
put(t, 1, 'Total', BOLD, fill=TOTF, border=True)
for c in (2, 3, 4, 5, 7, 8):
    put(t, c, '=SUM(%s%d:%s%d)' % (gl(c), first, gl(c), last), BOLD, MONEY,
        fill=TOTF, border=True)
put(t, 6, '=IF(D%d=0,"",E%d/D%d)' % (t, t, t), BOLD, PCT, fill=TOTF, border=True)

put(t + 2, 1, 'Variance ex GST is Xero less the BAS. A positive figure means Xero '
              'carries more income than was declared at G1.', NOTE)

# ------------------------------------------------------------- accrual block
a0 = t + 4
put(a0, 1, 'Accrual basis, for reference only', SECT)
bandrow(a0 + 1, 1, ['Quarter', 'Xero accrual\nex GST', 'Xero cash\nex GST',
                    'Accrual less cash', 'BAS G1\nex GST',
                    'Accrual less BAS'])
af, al = a0 + 2, a0 + 1 + len(QTRS)
for i, (name, bas, cash, accr, oth) in enumerate(QTRS):
    r = af + i
    src = first + i
    put(r, 1, name, BODY, border=True)
    put(r, 2, accr, INPUT, MONEY, border=True)
    put(r, 3, '=D%d' % src, BODY, MONEY, border=True)
    put(r, 4, '=B%d-C%d' % (r, r), BODY, MONEY, border=True)
    put(r, 5, '=C%d' % src, BODY, MONEY, border=True)
    put(r, 6, '=B%d-E%d' % (r, r), BODY, MONEY, border=True)
at = al + 1
put(at, 1, 'Total', BOLD, fill=TOTF, border=True)
for c in range(2, 7):
    put(at, c, '=SUM(%s%d:%s%d)' % (gl(c), af, gl(c), al), BOLD, MONEY,
        fill=TOTF, border=True)

put(at + 2, 1, 'The controller pack runs on accrual. These figures will not tie '
               'to any BAS and are here only so the two bases can be seen together.',
    NOTE)

# ---------------------------------------------------------- other income block
o0 = at + 4
put(o0, 1, 'Other income, excluded from G1', SECT)
bandrow(o0 + 1, 1, ['Quarter', 'Interest received\nand other income',
                    'Xero cash ex GST\nplus other income', 'BAS G1 ex GST',
                    'Variance if other income\nwere included'])
of_, ol = o0 + 2, o0 + 1 + len(QTRS)
for i, (name, bas, cash, accr, oth) in enumerate(QTRS):
    r = of_ + i
    src = first + i
    put(r, 1, name, BODY, border=True)
    put(r, 2, oth, INPUT, MONEY, border=True)
    put(r, 3, '=D%d+B%d' % (src, r), BODY, MONEY, border=True)
    put(r, 4, '=C%d' % src, BODY, MONEY, border=True)
    put(r, 5, '=C%d-D%d' % (r, r), BODY, MONEY, border=True)
ot = ol + 1
put(ot, 1, 'Total', BOLD, fill=TOTF, border=True)
for c in range(2, 6):
    put(ot, c, '=SUM(%s%d:%s%d)' % (gl(c), of_, gl(c), ol), BOLD, MONEY,
        fill=TOTF, border=True)

put(ot + 2, 1, 'Interest received is a financial supply and is not declared at G1. '
               'The cash basis block above ties without it, which is what says it '
               'was correctly left out.', NOTE)

# ------------------------------------------------------------------- findings
f0 = ot + 4
put(f0, 1, 'What this shows', SECT)
findings = [
    'The BAS is lodged on cash basis with G1 reported GST inclusive. '
    'On that basis two of the three quarters agree with Xero to within 400 dollars.',
    'The March 2026 quarter is the only one that moves. Xero carries more income '
    'than was declared, and it is the bulk of the total variance.',
    'GST free sales are not the explanation. On both lodged statements the GST at '
    '1A is exactly one eleventh of G1, so every sale declared was taxable. '
    'The March variance is a real difference, not a rounding or a mix effect.',
    'To settle the March quarter, run the Xero GST Reconciliation report or the '
    'Activity Statement report for that quarter. The profit and loss will not do it.',
    'These three quarters span two financial years. January to June 2026 is FY26, '
    'July to September 2026 is FY27. The October to December 2025 quarter is not here, '
    'so this is not a full year of either.',
]
for i, line in enumerate(findings):
    r = f0 + 1 + i
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    c = put(r, 1, line, BODY)
    c.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws.row_dimensions[r].height = 30

# ------------------------------------------------------------- attachments
k0 = f0 + 7
put(k0, 1, 'Lodged activity statements', SECT)
put(k0 + 1, 1, 'Double click an icon to open the statement. '
               'The April to June 2026 statement has not been supplied.', NOTE)
ws.merge_cells(start_row=k0 + 1, start_column=1, end_row=k0 + 1, end_column=8)
ws.cell(k0 + 1, 1).alignment = Alignment(horizontal='left', vertical='top',
                                         wrap_text=True)
for rr in range(k0 + 2, k0 + 11):
    ws.row_dimensions[rr].height = 15
ATTACH_ROW = k0 + 2

ws.column_dimensions['A'].width = 44
for c in range(2, 9):
    ws.column_dimensions[gl(c)].width = 17
ws.freeze_panes = 'A13'

for r in (5, 6, 7, t + 2, at + 2, ot + 2):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(r, 1).alignment = Alignment(horizontal='left', vertical='top',
                                        wrap_text=True)
    ws.row_dimensions[r].height = 15

ws.print_area = 'A1:H58'
ws.page_setup.orientation = 'landscape'
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True

wb.save(OUT)
print('wrote', OUT)
print('ATTACH_ROW', ATTACH_ROW)
