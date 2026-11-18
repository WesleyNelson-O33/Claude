"""Add a Sales to BAS tab: trading income per the P&L bridged to the lodged G1.

Starts from Danica's latest workbook so her own tabs and edits are kept. The
raw tabs are not touched. Every figure is a formula off the P&L tab or the GL
raw tab, except G1 as lodged, which only exists on the activity statement.
"""
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

SRC = ('/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a'
       '/scratchpad/bas2/w.xlsx')
OUT = 'recon/CTS BAS to Xero - Jan to Sep 2026.xlsx'

GLS = "'GL Jan - Sep 2026 ( Raw Data) '"
GSS = "'GST Jan - Sep 2026 (Raw Data)'"
PL = "'P&L'"
GST_FIRST, GST_LAST = 9, 1884

NAVY, WHITE, GREY, TEAL = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4', 'FF005358'
BLUE, RED = 'FF0070C0', 'FFC01C28'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTF = PatternFill('solid', start_color=GREY, end_color=GREY)
RULE = Side(style='thin', color='FFC9CED6')
HEAD = Font(name='Calibri', sz=11, b=True, color=WHITE)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
BODY = Font(name='Calibri', sz=11, color=NAVY)
INPUT = Font(name='Calibri', sz=11, color=BLUE)
WARN = Font(name='Calibri', sz=11, b=True, color=RED)
TITLE = Font(name='Calibri', sz=16, b=True, color=NAVY)
SUB = Font(name='Calibri', sz=11, color='FF6B7585')
SECT = Font(name='Calibri', sz=12, b=True, color=TEAL)
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')
MONEY = '#,##0.00;[RED](#,##0.00);"-"'

# quarter, P&L month columns, GL date range, G1 as lodged
QTRS = [('Jan - Mar 2026', ('H', 'I', 'J'), (2026, 1, 1), (2026, 3, 31), 1324425.00),
        ('Apr - Jun 2026', ('E', 'F', 'G'), (2026, 4, 1), (2026, 6, 30), 1678257.00),
        ('Jul - Sep 2026', ('B', 'C', 'D'), (2026, 7, 1), (2026, 9, 30), 1909366.00)]

# income account, first and last row of its section on the GL raw tab
SPANS = [
    ('Ad Hoc Service Labour', 111, 259),
    ('Conference Call Charge', 1531, 1575),
    ('Contract Support Staff', 1583, 1910),
    ('Equipment Hires', 8127, 8726),
    ('Equipment Sales', 8728, 8835),
    ('Freight & Travel - INTEGRATION', 8837, 8844),
    ('Freight & Travel - PRD', 8846, 8958),
    ('Freight & Travel - VID', 8960, 8963),
    ('Installation Labour', 11868, 11918),
    ('Meal & Travel - ONS', 12196, 12216),
    ('Meal & Travel - PRD', 12218, 12260),
    ('Mobile Phone Plans', 12262, 12387),
    ('Production Labour', 15340, 16099),
    ('Subscriptions & Licences - Income', 17246, 17477),
    ('Tech Consulting/Project Manage', 18408, 18447),
    ('Tech Event Management', 18449, 18514),
    ('Video Labour', 21915, 22075),
]

shutil.copy(SRC, OUT)
wb = openpyxl.load_workbook(OUT)
if 'Sales to BAS' in wb.sheetnames:
    del wb['Sales to BAS']
ws = wb.create_sheet('Sales to BAS', 2)
ws.sheet_view.showGridLines = False
COLS = ['B', 'C', 'D']


def put(r, c, v, font=BODY, fmt=None, fill=None, border=False, wrap=False):
    cell = ws.cell(r, c)
    cell.value = v
    cell.font = font
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if border:
        cell.border = Border(top=RULE, bottom=RULE, left=RULE, right=RULE)
    if wrap:
        cell.alignment = Alignment(horizontal='left', vertical='top',
                                   wrap_text=True)
    return cell


def band(r):
    for i, lab in enumerate(['', 'Jan - Mar 2026', 'Apr - Jun 2026',
                             'Jul - Sep 2026', 'Total']):
        c = put(r, 1 + i, lab, HEAD, fill=BAND)
        c.alignment = Alignment(horizontal='left' if i == 0 else 'right',
                                vertical='center', wrap_text=True)
    ws.row_dimensions[r].height = 42


def note(r, text, font=NOTE, h=28):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    put(r, 1, text, font, wrap=True)
    ws.row_dimensions[r].height = h


def taxable(span, q):
    """Credits less debits on GST on Income lines in one account, one quarter."""
    name, a, b = span
    s, e = QTRS[q][2], QTRS[q][3]
    rng = lambda col: '%s!$%s$%d:$%s$%d' % (GLS, col, a, col, b)
    crit = (',%s,"GST on Income",%s,">="&DATE(%d,%d,%d),%s,"<="&DATE(%d,%d,%d)'
            % (rng('J'), rng('A'), s[0], s[1], s[2], rng('A'), e[0], e[1], e[2]))
    return '=SUMIFS(%s%s)-SUMIFS(%s%s)' % (rng('F'), crit, rng('E'), crit)


put(1, 1, 'Sales reconciled to the lodged BAS', TITLE)
put(2, 1, 'Corporate Technology Services Pty Ltd. January to September 2026.', SUB)
note(4, 'For anyone reconciling the BAS back to sales. Trading income on the '
        'P&L is not the same as sales declared at G1, for two reasons, and this '
        'tab shows both of them with the amounts.', SUB, 30)

r = 6
put(r, 1, 'Taxable sales by account, from the general ledger', SECT)
r += 1
band(r)
FIRST = r + 1
for i, span in enumerate(SPANS):
    rr = FIRST + i
    put(rr, 1, '%s   (GL rows %d to %d)' % (span[0], span[1], span[2]), BODY,
        border=True)
    for j in range(3):
        put(rr, 2 + j, taxable(span, j), BODY, MONEY, border=True)
    put(rr, 5, '=SUM(B%d:D%d)' % (rr, rr), BODY, MONEY, border=True)
LAST = FIRST + len(SPANS) - 1
TAX = LAST + 1
put(TAX, 1, 'Taxable sales, excluding GST', BOLD, fill=TOTF, border=True)
for c in range(2, 6):
    put(TAX, c, '=SUM(%s%d:%s%d)' % (gl(c), FIRST, gl(c), LAST), BOLD, MONEY,
        fill=TOTF, border=True)
note(TAX + 1, 'Only lines the ledger marks GST on Income. The row range after '
              'each account is its section on the GL raw tab.')

r = TAX + 3
put(r, 1, 'From sales to G1', SECT)
r += 1
band(r)
B = r + 1
pl_sum = lambda cols: '=' + '+'.join("%s!$%s$26" % (PL, c) for c in cols)
rows = [
    ('Trading income per the P&L', [pl_sum(q[1]) for q in QTRS]),
    ('Less revenue carrying no GST', ['=%s%d-%s%d' % (c, B, c, B + 2)
                                      for c in COLS]),
    ('Taxable sales, excluding GST', ['=%s%d' % (c, TAX) for c in COLS]),
    ('GST on those sales, at one tenth', ['=%s%d/10' % (c, B + 2) for c in COLS]),
    ('G1 on an accrual basis', ['=%s%d+%s%d' % (c, B + 2, c, B + 3)
                                for c in COLS]),
    ('G1 as lodged, cash basis', [q[4] for q in QTRS]),
    ('Difference, timing', ['=%s%d-%s%d' % (c, B + 4, c, B + 5) for c in COLS]),
]
for i, (lab, vals) in enumerate(rows):
    rr = B + i
    strong = lab.startswith(('Taxable sales', 'G1 on', 'Difference'))
    put(rr, 1, lab, BOLD if strong else BODY,
        fill=TOTF if strong else None, border=True)
    for j, v in enumerate(vals):
        put(rr, 2 + j, v, INPUT if lab.startswith('G1 as lodged') else
            (BOLD if strong else BODY), MONEY,
            fill=TOTF if strong else None, border=True)
    put(rr, 5, '=SUM(B%d:D%d)' % (rr, rr), BOLD if strong else BODY, MONEY,
        fill=TOTF if strong else None, border=True)

r = B + len(rows) + 1
put(r, 1, 'The two reasons sales do not equal G1', WARN)
note(r + 1, '1. Revenue carrying no GST. Mostly the subscription deferral '
            'journals on Subscriptions and Licences Income, which move revenue '
            'between months without GST because the GST went on the original '
            'invoice. They are correctly coded BAS Excluded. They sit in '
            'trading income but never reach G1.', BODY, 48)
note(r + 2, '2. Timing. The BAS is lodged on a cash basis, so a sale reaches G1 '
            'in the quarter the customer pays, not the quarter it is invoiced. '
            'CTS raises large invoices in the last days of a month, so the GST '
            'on them lands on the next BAS.', BODY, 48)
note(r + 3, 'Neither is an error and no GST is missed. The first is a '
            'presentation difference, the second reverses in the following '
            'quarter.', BODY, 30)

ws.column_dimensions['A'].width = 48
for c in range(2, 6):
    ws.column_dimensions[gl(c)].width = 17
ws.print_area = 'A1:E%d' % (r + 4)
ws.page_setup.orientation = 'landscape'
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True

# --------------------------------------------------- trim the BAS to Xero tab
# The "Other income, excluded from G1" block was a working check from when the
# lodgement basis was still being proved. The hypothetical "variance if other
# income were included" column answers a question nobody asked, so it goes.
bx = wb['BAS to Xero']
for rr in range(21, 27):
    for cc in range(1, 9):
        c = bx.cell(rr, cc)
        c.value = None
        c.fill = PatternFill(fill_type=None)
        c.border = Border()
bx.merge_cells(start_row=21, start_column=1, end_row=21, end_column=8)
n = bx.cell(21, 1)
n.value = ('Interest received is a financial supply and is not declared at G1. '
           'It is excluded from every figure above. The amounts are on the BAS '
           'Calc tab.')
n.font = NOTE
n.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
bx.row_dimensions[21].height = 16

wb.save(OUT)
print('wrote', OUT, '| TAX row', TAX, '| bridge at', B)
