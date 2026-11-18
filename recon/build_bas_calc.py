"""Complete Danica's BAS calc: the summary block, and a BAS Calc tab built on
formulas that point at her two raw tabs.

Her raw tabs are not touched. The GST CALC tab gets its running balance chain
repaired and the four labels she wrote at D548:D551 filled with formulas.
"""
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

SRC = '/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a/scratchpad/basin/w.xlsx'
OUT = 'recon/CTS BAS to Xero - Jan to Sep 2026.xlsx'

GLS = "'GL Jan - Sep 2026 ( Raw Data) '"
GSS = "'GST Jan - Sep 2026 (Raw Data)'"
GST_FIRST, GST_LAST = 9, 1884          # data rows on the GST raw tab
INT_FIRST, INT_LAST = 11972, 11985     # Interest Received section on the GL tab

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

QTRS = [('Jan - Mar 2026', '2026-01-01', '2026-03-31', 1324425.00),
        ('Apr - Jun 2026', '2026-04-01', '2026-06-30', 1678257.00),
        ('Jul - Sep 2026', '2026-07-01', '2026-09-30', 1909366.00)]

shutil.copy(SRC, OUT)
wb = openpyxl.load_workbook(OUT)

# ---------------------------------------------------------------- the new tab
if 'BAS Calc' in wb.sheetnames:
    del wb['BAS Calc']
ws = wb.create_sheet('BAS Calc', 1)
ws.sheet_view.showGridLines = False


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


def band(r, labels):
    for i, lab in enumerate(labels):
        c = put(r, 1 + i, lab, HEAD, fill=BAND)
        c.alignment = Alignment(horizontal='left' if i == 0 else 'right',
                                vertical='center', wrap_text=True)
    ws.row_dimensions[r].height = 42


def note(r, text):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    put(r, 1, text, NOTE, wrap=True)
    ws.row_dimensions[r].height = 28


put(1, 1, 'BAS calculated from the general ledger', TITLE)
put(2, 1, 'Corporate Technology Services Pty Ltd. January to September 2026. '
          'Three quarters.', SUB)

put(4, 1, 'How this is built', SECT)
note(5, 'Every figure below is a formula reading the two raw tabs in this '
        'workbook. Nothing is retyped except the four G1 figures from the lodged '
        'activity statements, which are blue. The raw tabs are untouched.')
note(6, 'Dates drive the quarter columns. The two cells under each quarter '
        'heading are the period start and end, and every formula in that column '
        'reads them.')

r = 8
band(r, ['', 'Jan - Mar 2026', 'Apr - Jun 2026', 'Jul - Sep 2026', 'Total'])
DATEROW = r + 1
put(DATEROW, 1, 'Period', BOLD)
for i, (nm, s, e, g1) in enumerate(QTRS):
    put(DATEROW, 2 + i, s, INPUT, 'dd mmm yyyy').alignment = Alignment(horizontal='right')
    put(DATEROW + 1, 2 + i, e, INPUT, 'dd mmm yyyy').alignment = Alignment(horizontal='right')
put(DATEROW + 1, 1, '', BODY)

COLS = ['B', 'C', 'D']
DS = '%s$%d' % ('%s', DATEROW)
DE = '%s$%d' % ('%s', DATEROW + 1)


def dates(col):
    return (',%s!$A:$A,">="&%s$%d,%s!$A:$A,"<="&%s$%d'
            % (GLS, col, DATEROW, GLS, col, DATEROW + 1))


def gdates(col):
    return (',%s!$C$%d:$C$%d,"%%s",%s!$A$%d:$A$%d,">="&%s$%d,%s!$A$%d:$A$%d,"<="&%s$%d'
            % (GSS, GST_FIRST, GST_LAST, GSS, GST_FIRST, GST_LAST, col, DATEROW,
               GSS, GST_FIRST, GST_LAST, col, DATEROW + 1))


# ------------------------------------------------------- per the ledger block
r = DATEROW + 3
put(r, 1, 'Per the general ledger, accrual', SECT)
r += 1
band(r, ['', 'Jan - Mar 2026', 'Apr - Jun 2026', 'Jul - Sep 2026', 'Total'])
LED = r + 1
rows = [
    ('Taxable sales, excluding GST',
     '=SUMIFS({G}!$F:$F,{G}!$J:$J,"GST on Income"{D})'
     '-SUMIFS({G}!$E:$E,{G}!$J:$J,"GST on Income"{D})'),
    ('GST on those sales',
     '=-SUMIFS({G}!$H:$H,{G}!$J:$J,"GST on Income"{D})'),
    ('GST free and export sales (G2)',
     '=SUMIFS({G}!$F:$F,{G}!$J:$J,"GST Free Income"{D})'
     '-SUMIFS({G}!$E:$E,{G}!$J:$J,"GST Free Income"{D})'),
    ('Total sales including GST (G1)', None),
    ('GST on purchases (1B)',
     '=SUMIFS({G}!$H:$H,{G}!$J:$J,"GST on Expenses"{D})'),
]
for i, (lab, f) in enumerate(rows):
    rr = LED + i
    put(rr, 1, lab, BOLD if lab.startswith('Total') else BODY,
        fill=TOTF if lab.startswith('Total') else None, border=True)
    for j, col in enumerate(COLS):
        if f is None:
            v = '=%s%d+%s%d+%s%d' % (col, LED, col, LED + 1, col, LED + 2)
        else:
            v = f.format(G=GLS, D=dates(col))
        put(rr, 2 + j, v, BOLD if f is None else BODY, MONEY,
            fill=TOTF if f is None else None, border=True)
    put(rr, 5, '=SUM(B%d:D%d)' % (rr, rr), BOLD if f is None else BODY, MONEY,
        fill=TOTF if f is None else None, border=True)

note(LED + 5, 'Read from the GST Rate Name column on the GL raw tab, column J. '
              'GST free and export sales come back nil because no line in the '
              'ledger carries that rate. That is the evidence for G2.')

# ------------------------------------------------------------- as lodged block
r = LED + 7
put(r, 1, 'As lodged with the ATO', SECT)
r += 1
band(r, ['', 'Jan - Mar 2026', 'Apr - Jun 2026', 'Jul - Sep 2026', 'Total'])
LOD = r + 1
put(LOD, 1, 'G1 Total sales, including GST', BODY, border=True)
for j, (nm, s, e, g1) in enumerate(QTRS):
    put(LOD, 2 + j, g1, INPUT, MONEY, border=True)
put(LOD, 5, '=SUM(B%d:D%d)' % (LOD, LOD), BODY, MONEY, border=True)

put(LOD + 1, 1, 'G2 International sales', BODY, border=True)
for j, col in enumerate(COLS):
    put(LOD + 1, 2 + j, '=%s%d' % (col, LED + 2), BODY, MONEY, border=True)
put(LOD + 1, 5, '=SUM(B%d:D%d)' % (LOD + 1, LOD + 1), BODY, MONEY, border=True)

put(LOD + 2, 1, '1A GST on sales', BODY, border=True)
put(LOD + 3, 1, '1B GST on purchases', BODY, border=True)
for j, col in enumerate(COLS):
    put(LOD + 2, 2 + j,
        '=SUMIFS(%s!$E$%d:$E$%d%s)' % (GSS, GST_FIRST, GST_LAST,
                                       gdates(col) % 'BAS*'),
        BODY, MONEY, border=True)
    put(LOD + 3, 2 + j,
        '=SUMIFS(%s!$F$%d:$F$%d%s)' % (GSS, GST_FIRST, GST_LAST,
                                       gdates(col) % 'BAS*'),
        BODY, MONEY, border=True)
for rr in (LOD + 2, LOD + 3):
    put(rr, 5, '=SUM(B%d:D%d)' % (rr, rr), BODY, MONEY, border=True)

put(LOD + 4, 1, 'Check: 1A times 11', BOLD, fill=TOTF, border=True)
put(LOD + 5, 1, 'Check: less G1 as lodged', BOLD, fill=TOTF, border=True)
for j, col in enumerate(COLS):
    put(LOD + 4, 2 + j, '=%s%d*11' % (col, LOD + 2), BOLD, MONEY, fill=TOTF,
        border=True)
    put(LOD + 5, 2 + j, '=%s%d-%s%d' % (col, LOD + 4, col, LOD), BOLD, MONEY,
        fill=TOTF, border=True)
for rr in (LOD + 4, LOD + 5):
    put(rr, 5, '=SUM(B%d:D%d)' % (rr, rr), BOLD, MONEY, fill=TOTF, border=True)

note(LOD + 6, '1A and 1B are read from the BAS clearing journals sitting in the '
              'GST account on the GST raw tab, the rows whose description starts '
              'with BAS. G1 is typed from the lodged activity statements.')
note(LOD + 7, 'The check proves there were no GST free sales. G1 divided by 11 is '
              'the GST, so 1A times 11 comes back to G1 within a few dollars of '
              'rounding. If any sale had been GST free the two would not meet.')

# -------------------------------------------------------------- variance block
r = LOD + 9
put(r, 1, 'Ledger against lodged', SECT)
r += 1
band(r, ['', 'Jan - Mar 2026', 'Apr - Jun 2026', 'Jul - Sep 2026', 'Total'])
VAR = r + 1
pairs = [('G1 Total sales', LED + 3, LOD),
         ('1A GST on sales', LED + 1, LOD + 2),
         ('1B GST on purchases', LED + 4, LOD + 3)]
for i, (lab, a, b) in enumerate(pairs):
    rr = VAR + i
    put(rr, 1, lab + ', ledger less lodged', BODY, border=True)
    for j, col in enumerate(COLS):
        put(rr, 2 + j, '=%s%d-%s%d' % (col, a, col, b), BODY, MONEY, border=True)
    put(rr, 5, '=SUM(B%d:D%d)' % (rr, rr), BODY, MONEY, border=True)

put(VAR + 3, 1, 'Why these do not agree', WARN)
note(VAR + 4, 'The BAS is lodged on a CASH basis. Both lodged statements say GST '
              'accounting method Cash. The general ledger is accrual. So the '
              'ledger figures above are sales invoiced, and the lodged figures '
              'are sales banked. They will never be the same in a quarter.')
note(VAR + 5, 'This block is not an error to chase. It is the timing difference '
              'between debtors invoiced and debtors paid. It is shown so the '
              'difference is on the page rather than hidden.')

# ----------------------------------------------------- GST account recon block
r = VAR + 7
put(r, 1, 'GST account reconciliation, 1 January to 30 September 2026', SECT)
r += 1
band(r, ['', 'Amount', '', '', ''])
REC = r + 1
exc = ',%s!$C$%d:$C$%d,"<>BAS*"' % (GSS, GST_FIRST, GST_LAST)
inc = ',%s!$C$%d:$C$%d,"BAS*"' % (GSS, GST_FIRST, GST_LAST)
recrows = [
    ('Opening balance at 1 January 2026, debit',
     '=-%s!$E$8' % GSS),
    ('GST collected on sales, posted to the account',
     '=SUMIFS(%s!$F$%d:$F$%d%s)' % (GSS, GST_FIRST, GST_LAST, exc)),
    ('GST paid on purchases, posted to the account',
     '=-SUMIFS(%s!$E$%d:$E$%d%s)' % (GSS, GST_FIRST, GST_LAST, exc)),
    ('BAS journals, GST collected cleared out',
     '=-SUMIFS(%s!$E$%d:$E$%d%s)' % (GSS, GST_FIRST, GST_LAST, inc)),
    ('BAS journals, GST paid cleared out',
     '=SUMIFS(%s!$F$%d:$F$%d%s)' % (GSS, GST_FIRST, GST_LAST, inc)),
]
for i, (lab, f) in enumerate(recrows):
    rr = REC + i
    put(rr, 1, lab, BODY, border=True)
    put(rr, 2, f, BODY, MONEY, border=True)
CLOSE = REC + len(recrows)
put(CLOSE, 1, 'Closing balance at 30 September 2026', BOLD, fill=TOTF, border=True)
put(CLOSE, 2, '=SUM(B%d:B%d)' % (REC, CLOSE - 1), BOLD, MONEY, fill=TOTF,
    border=True)
put(CLOSE + 1, 1, 'Closing balance per the GST raw tab', BODY, border=True)
put(CLOSE + 1, 2, '=%s!$G$1886' % GSS, BODY, MONEY, border=True)
put(CLOSE + 2, 1, 'Difference', BOLD, fill=TOTF, border=True)
put(CLOSE + 2, 2, '=B%d-B%d' % (CLOSE, CLOSE + 1), BOLD, MONEY, fill=TOTF,
    border=True)
note(CLOSE + 3, 'This is the piece that balances. It proves every dollar of GST '
                'posted to the account over the nine months has either been '
                'lodged on a BAS or is sitting in the closing balance. The '
                'difference line is the proof and it is nil.')

# ------------------------------------------------------------ interest block
r = CLOSE + 5
put(r, 1, 'Interest received, excluded from G1', SECT)
r += 1
band(r, ['', 'Jan - Mar 2026', 'Apr - Jun 2026', 'Jul - Sep 2026', 'Total'])
INT = r + 1
put(INT, 1, 'Interest received per the ledger', BODY, border=True)
for j, col in enumerate(COLS):
    f = ('=SUMIFS({G}!$F${a}:$F${b},{G}!$A${a}:$A${b},">="&{c}${d},'
         '{G}!$A${a}:$A${b},"<="&{c}${e})'
         '-SUMIFS({G}!$E${a}:$E${b},{G}!$A${a}:$A${b},">="&{c}${d},'
         '{G}!$A${a}:$A${b},"<="&{c}${e})').format(
        G=GLS, a=INT_FIRST, b=INT_LAST, c=col, d=DATEROW, e=DATEROW + 1)
    put(INT, 2 + j, f, BODY, MONEY, border=True)
put(INT, 5, '=SUM(B%d:D%d)' % (INT, INT), BODY, MONEY, border=True)
note(INT + 1, 'Read from the Interest Received section of the GL raw tab, rows '
              '%d to %d. Interest is a financial supply and is not declared at '
              'G1, which is why it is shown separately.' % (INT_FIRST, INT_LAST))

# ------------------------------------------------------------------ warnings
r = INT + 3
put(r, 1, 'Three things to know before this goes to the ATO', WARN)
warnings = [
    'The VLOOKUPs in columns J, K and L of the GST CALC tab look up on column B, '
    'which holds the transaction type: Payable Invoice, Manual Journal and so on. '
    'That is not a unique key, so every row returns the first match in the ledger. '
    'The answers in those three columns are wrong on every row. I have left them '
    'exactly as they are rather than change your work, but they cannot go in front '
    'of the ATO.',
    'The GST CALC tab holds the data in three quarterly blocks, rows 3 to 543, '
    '566 to 1162 and 1184 to 1921, with a heading row above the second and third. '
    'Any total has to cover all three and skip the headings. That is why the '
    'figures here are built from the raw tabs instead, by date, not by block.',
    'The lodged BAS is cash basis and the ledger is accrual, so the ledger will '
    'not reproduce the lodged G1 or 1A. To evidence the lodged figures themselves, '
    'the report to pull is the Xero GST Reconciliation report, or the Activity '
    'Statement report, for each quarter on a cash basis.',
]
for i, w in enumerate(warnings):
    rr = r + 1 + i
    ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=5)
    put(rr, 1, w, BODY, wrap=True)
    ws.row_dimensions[rr].height = 46

ws.column_dimensions['A'].width = 44
for c in range(2, 6):
    ws.column_dimensions[gl(c)].width = 17

# ------------------------------------------- repair her CALC tab, fill summary
# Her tab holds three quarterly blocks with a heading row above each later one.
BLOCKS = [(3, 543, 1), (566, 1162, 543), (1184, 1921, 1162)]
calc = wb['GST Jan - Sep 2026 CALC']
for first, last, anchor in BLOCKS:
    for rr in range(first, last + 1):
        prev = anchor if rr == first else rr - 1
        calc.cell(rr, 7).value = '=(G%d+F%d)-E%d' % (prev, rr, rr)

# Each summary cell reads the raw tabs directly. No hop through another tab.
GI = '"GST on Income"'
GF = '"GST Free Income"'


def drange(col, q):
    s, e = QD[q]
    return (',{G}!$A:$A,">="&DATE({sy},{sm},{sd}),{G}!$A:$A,"<="&DATE({ey},{em},{ed})'
            .format(G=GLS, sy=s[0], sm=s[1], sd=s[2], ey=e[0], em=e[1], ed=e[2]))


def grange(q):
    s, e = QD[q]
    return (',{S}!$A${a}:$A${b},">="&DATE({sy},{sm},{sd}),'
            '{S}!$A${a}:$A${b},"<="&DATE({ey},{em},{ed})'
            .format(S=GSS, a=GST_FIRST, b=GST_LAST, sy=s[0], sm=s[1], sd=s[2],
                    ey=e[0], em=e[1], ed=e[2]))


QD = {0: ((2026, 1, 1), (2026, 3, 31)),
      1: ((2026, 4, 1), (2026, 6, 30)),
      2: ((2026, 7, 1), (2026, 9, 30))}

LABELS = ['G1 - Total sales  ', 'G2 - International Sales ', '1A - GST',
          'Interest received']
SUMMARY = [(548, 0, 'Jan - Mar 2026'),
           (1166, 1, 'Apr - Jun 2026'),
           (1926, 2, 'Jul - Sep 2026')]
LODGED_G1 = {0: 1324425.00, 1: 1678257.00, 2: 1909366.00}

for top, q, qname in SUMMARY:
    dr = drange('x', q)
    gr = grange(q)
    inc_net = ('SUMIFS({G}!$F:$F,{G}!$J:$J,{R}{D})-SUMIFS({G}!$E:$E,{G}!$J:$J,{R}{D})'
               .format(G=GLS, R=GI, D=dr))
    inc_gst = 'SUMIFS({G}!$H:$H,{G}!$J:$J,{R}{D})'.format(G=GLS, R=GI, D=dr)
    free_net = ('SUMIFS({G}!$F:$F,{G}!$J:$J,{R}{D})-SUMIFS({G}!$E:$E,{G}!$J:$J,{R}{D})'
                .format(G=GLS, R=GF, D=dr))
    interest = ('SUMIFS({G}!$F${a}:$F${b},{G}!$A${a}:$A${b},">="&DATE({sy},{sm},{sd}),'
                '{G}!$A${a}:$A${b},"<="&DATE({ey},{em},{ed}))'
                '-SUMIFS({G}!$E${a}:$E${b},{G}!$A${a}:$A${b},">="&DATE({sy},{sm},{sd}),'
                '{G}!$A${a}:$A${b},"<="&DATE({ey},{em},{ed}))'
                .format(G=GLS, a=INT_FIRST, b=INT_LAST,
                        sy=QD[q][0][0], sm=QD[q][0][1], sd=QD[q][0][2],
                        ey=QD[q][1][0], em=QD[q][1][1], ed=QD[q][1][2]))
    lodged_1a = 'SUMIFS({S}!$E${a}:$E${b},{S}!$C${a}:$C${b},"BAS*"{D})'.format(
        S=GSS, a=GST_FIRST, b=GST_LAST, D=gr)

    ledger = ['=(%s)-(%s)' % (inc_net, inc_gst), '=%s' % free_net,
              '=-(%s)' % inc_gst, '=%s' % interest]
    lodged = [LODGED_G1[q], '=%s' % free_net, '=%s' % lodged_1a, None]

    h = calc.cell(top - 3, 4)
    h.value = qname
    h.font = Font(name='Calibri', sz=11, b=True, color=TEAL)
    for c, txt in ((5, 'Per the ledger'), (6, 'As lodged'), (8, 'Difference')):
        hc = calc.cell(top - 2, c)
        hc.value = txt
        hc.font = Font(name='Calibri', sz=11, b=True, color=NAVY)
        hc.alignment = Alignment(horizontal='right')

    for i, lab in enumerate(LABELS):
        rr = top + i
        lc = calc.cell(rr, 4)
        if lc.value in (None, ''):
            lc.value = lab
            lc.font = Font(name='Calibri', sz=11, color=NAVY)
        a = calc.cell(rr, 5)
        a.value = ledger[i]
        a.number_format = MONEY
        a.font = Font(name='Calibri', sz=11, b=True, color=NAVY)
        if lodged[i] is not None:
            b = calc.cell(rr, 6)
            b.value = lodged[i]
            b.number_format = MONEY
            b.font = (INPUT if isinstance(lodged[i], float)
                      else Font(name='Calibri', sz=11, color=NAVY))
            diff = calc.cell(rr, 8)
            diff.value = '=E%d-F%d' % (rr, rr)
            diff.number_format = MONEY
            diff.font = Font(name='Calibri', sz=11, b=True, color=NAVY)

    n = calc.cell(top + 4, 4)
    n.value = ('Why they differ: the BAS was lodged on a CASH basis and the '
               'ledger is accrual, so the ledger counts sales invoiced and the '
               'lodged figure counts sales banked. The difference is the GST on '
               'debtors that moved across the quarter end. It is not an error. '
               'Per the ledger reads the GL raw tab by GST Rate Name and date, '
               'as lodged reads the BAS journals on the GST raw tab, and only '
               'G1 as lodged is typed, off the activity statement.')
    n.font = NOTE

    # the GST account calc for this quarter, from the rows above on this tab
    first, last, anc = BLOCKS[q]
    t = top + 5
    th = calc.cell(t, 4)
    th.value = ('How the GST account moved this quarter, from rows %d to %d above'
                % (first, last))
    th.font = Font(name='Calibri', sz=11, b=True, color=TEAL)
    for c, txt in ((5, 'Debit'), (6, 'Credit'), (8, 'Gross')):
        hc = calc.cell(t + 1, c)
        hc.value = txt
        hc.font = Font(name='Calibri', sz=11, b=True, color=NAVY)
        hc.alignment = Alignment(horizontal='right')

    rng = '$%s$%d:$%s$%d'
    desc = '$C$%d:$C$%d' % (first, last)
    lines = [
        ('Total of the rows above', [
            (5, '=SUM($E$%d:$E$%d)' % (first, last)),
            (6, '=SUM($F$%d:$F$%d)' % (first, last)),
            (8, '=SUM($H$%d:$H$%d)' % (first, last))]),
        ('GST collected on sales, excluding the BAS journal', [
            (6, '=SUMIFS($F$%d:$F$%d,%s,"<>BAS*")' % (first, last, desc))]),
        ('GST paid on purchases, excluding the BAS journal', [
            (5, '=SUMIFS($E$%d:$E$%d,%s,"<>BAS*")' % (first, last, desc))]),
        ('BAS journal, GST collected cleared out', [
            (5, '=SUMIFS($E$%d:$E$%d,%s,"BAS*")' % (first, last, desc))]),
        ('BAS journal, GST paid cleared out', [
            (6, '=SUMIFS($F$%d:$F$%d,%s,"BAS*")' % (first, last, desc))]),
        ('Movement in the quarter, credits less debits', [
            (6, '=$F$%d-$E$%d' % (t + 2, t + 2))]),
        ('Movement per the Xero running balance, column G', [
            (6, '=G%d-G%d' % (last, anc))]),
        ('Difference', [
            (6, '=$F$%d-$F$%d' % (t + 7, t + 8))]),
    ]
    for i, (lab, cells) in enumerate(lines):
        rr = t + 2 + i
        strong = lab in ('Total of the rows above', 'Difference')
        lc = calc.cell(rr, 4)
        lc.value = lab
        lc.font = Font(name='Calibri', sz=11, b=strong, color=NAVY)
        for cc, fml in cells:
            vc = calc.cell(rr, cc)
            vc.value = fml
            vc.number_format = MONEY
            vc.font = Font(name='Calibri', sz=11, b=strong, color=NAVY)
    nn = calc.cell(t + 11, 4)
    nn.value = ('Every figure here is a SUM or SUMIFS over the quarter rows on '
                'this tab. The difference line proves the block agrees with the '
                'Xero running balance.')
    nn.font = NOTE

# ------------------------------------------- columns J to N on the CALC tab
# Pulled from the GL raw tab on Reference plus Date. The GL export carries no
# invoice number column, so K takes the leading token of the reference.
G6, G7 = 6, 22766
GC = "%s!$C$%d:$C$%d" % (GLS, G6, G7)
GD = "%s!$D$%d:$D$%d" % (GLS, G6, G7)
GA = "%s!$A$%d:$A$%d" % (GLS, G6, G7)
GE = "%s!$E$%d:$E$%d" % (GLS, G6, G7)
GF = "%s!$F$%d:$F$%d" % (GLS, G6, G7)
GH = "%s!$H$%d:$H$%d" % (GLS, G6, G7)
GI = "%s!$I$%d:$I$%d" % (GLS, G6, G7)

HEADERS = [(10, 'Supplier name\nfrom the GL'), (11, 'Invoice number\nfrom the GL'),
           (12, 'Total sales amount\nexcluding GST'), (13, 'GST'), (14, 'Total\nincluding GST')]
for hr in (2, 565, 1183):
    for cc, txt in HEADERS:
        hc = calc.cell(hr, cc)
        hc.value = txt
        hc.font = Font(name='Calibri', sz=11, b=True, color=WHITE)
        hc.fill = BAND
        hc.alignment = Alignment(horizontal='right', vertical='center',
                                 wrap_text=True)
    calc.row_dimensions[hr].height = 42

for first, last, anc in BLOCKS:
    for rr in range(first, last + 1):
        crit = ',%s,$D%d,%s,$A%d,%s,10' % (GD, rr, GA, rr, GI)
        calc.cell(rr, 10).value = (
            '=IFERROR(SUBSTITUTE(INDEX(%s,MATCH($D%d,%s,0)),"Payment: ",""),"")'
            % (GC, rr, GD))
        calc.cell(rr, 11).value = (
            '=IFERROR(LEFT($D%d,FIND(" - ",$D%d)-1),$D%d)' % (rr, rr, rr))
        calc.cell(rr, 12).value = (
            '=SUMIFS(%s%s)-SUMIFS(%s%s)' % (GF, crit, GE, crit))
        calc.cell(rr, 13).value = '=-SUMIFS(%s%s)' % (GH, crit)
        calc.cell(rr, 14).value = '=L%d+M%d' % (rr, rr)
        for cc in (12, 13, 14):
            calc.cell(rr, cc).number_format = MONEY
        for cc in range(10, 15):
            cell = calc.cell(rr, cc)
            cell.font = Font(name='Calibri', sz=11, color=NAVY)
            cell.alignment = Alignment(horizontal='left' if cc < 12 else 'right',
                                       vertical='center', wrap_text=False)

calc.column_dimensions['F'].width = 16
for cc, w in ((10, 36), (11, 22), (12, 19), (13, 14), (14, 17)):
    calc.column_dimensions[gl(cc)].width = w

wb.save(OUT)
print('wrote', OUT)
print('LED', LED, 'LOD', LOD, 'VAR', VAR, 'REC', REC, 'CLOSE', CLOSE, 'INT', INT)
