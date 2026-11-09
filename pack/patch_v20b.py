"""Top 10 Clients, the aged receivables dump, and the DSO sheet."""
import os, json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = f'{HERE}/CTS Financial Controller Pack - v20.xlsx'
B = json.load(open(f'{HERE}/data/brand.json'))
NS = json.load(open(f'{HERE}/data/newsheets.json'))
AR = json.load(open(f'{HERE}/data/ar_sep26.json'))

TEAL, AQUA, BLUE, INK, WHITE = (B['dark_teal'], B['bright_teal'], B['primary_blue'],
                                B['charcoal'], B['white'])
RED, GREY = B['accent_red'], B['grey']
MONEY = '\\$#,##0;[Red]"($"#,##0\\);\\-'
MONEY2 = '\\$#,##0.00;[Red]"($"#,##0.00\\);\\-'
PCT = '0.0%;[Red](0.0%);\\-'
DAYS = '#,##0.0" days"'

CA, CB, CE, CJ = ('Cleanup!$A$5:$A$20004', 'Cleanup!$B$5:$B$20004',
                  'Cleanup!$E$5:$E$20004', 'Cleanup!$J$5:$J$20004')
FY0, MON = 'DATE(Setup!$B$7-1,Setup!$B$6,1)', 'Setup!$B$5'

wb = openpyxl.load_workbook(OUT)

def sheet(name, after):
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name, wb.sheetnames.index(after) + 1)
    ws.sheet_properties.tabColor = TEAL
    ws.sheet_view.showGridLines = False
    return ws

def title(c, t, sz=14):
    c.value = t; c.font = Font(name='Calibri', sz=sz, bold=True, color=TEAL)

def note(c, t):
    c.value = t; c.font = Font(name='Calibri', sz=9, italic=True, color=INK)

def head(ws, r, c0, labels, h=22):
    for i, t in enumerate(labels):
        c = ws.cell(r, c0 + i, t)
        c.font = Font(name='Calibri', sz=10, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=TEAL)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[r].height = h

def banner(ws, r, c0, c1, text, colour=TEAL, ink=WHITE):
    for c in range(c0, c1 + 1):
        x = ws.cell(r, c)
        x.fill = PatternFill('solid', fgColor=colour)
        x.font = Font(name='Calibri', sz=10, bold=True, color=ink)
    ws.cell(r, c0, text)
    ws.row_dimensions[r].height = 20

def body(ws, r, c0, c1, fmt=MONEY, bold=False):
    for c in range(c0, c1 + 1):
        x = ws.cell(r, c)
        x.number_format = fmt
        x.font = Font(name='Calibri', sz=10, bold=bold, color=INK)

def rule(ws, r, c0, c1):
    for c in range(c0, c1 + 1):
        ws.cell(r, c).border = Border(top=Side('thin', color=TEAL))
        ws.cell(r, c).font = Font(name='Calibri', sz=10, bold=True, color=TEAL)

def inp(ws, cell, val, fmt='General'):
    c = ws[cell]; c.value = val; c.number_format = fmt
    c.fill = PatternFill('solid', fgColor=AQUA)
    c.font = Font(name='Calibri', sz=11, bold=True, color=TEAL)
    c.alignment = Alignment(horizontal='center')

# ====================================================== Top 10 Clients =====
ws = sheet('Top 10 Clients', 'P&L by Dept')
W0 = 24
CUST = NS['customers']
W1 = W0 + max(len(CUST), 80) - 1
RNG = lambda col: f'${col}${W0}:${col}${W1}'

title(ws['A1'], 'Top 10 Clients')
note(ws['A2'], 'Revenue by customer, off the contact on every income line in the GL. The top ten '
               'pick themselves from the workings below, so a new customer appears on its own as '
               'soon as it earns. Everything here is a formula.')
ws['A3'] = '=TEXT(Setup!$B$5,"mmm yyyy")&"   |   year to date from "&TEXT(' + FY0 + ',"mmm yyyy")'
ws['A3'].font = Font(name='Calibri', sz=11, bold=True, color=TEAL)

banner(ws, 5, 1, 5, 'THE TOP TEN, YEAR TO DATE')
head(ws, 6, 1, ['Rank', 'Customer', 'This month', 'Year to date', 'Share of YTD'])
for n in range(1, 11):
    r = 6 + n
    ws.cell(r, 1, n).alignment = Alignment(horizontal='center')
    key = f'MATCH(LARGE({RNG("D")},$A{r}),{RNG("D")},0)'
    ws.cell(r, 2, f'=IFERROR(INDEX({RNG("A")},{key}),"")')
    ws.cell(r, 3, f'=IFERROR(INDEX({RNG("B")},{key}),0)')
    ws.cell(r, 4, f'=IFERROR(INDEX({RNG("C")},{key}),0)')
    ws.cell(r, 5, f"=IFERROR($D{r}/'P&L FY27'!$O$6,0)")
    body(ws, r, 3, 4)
    body(ws, r, 5, 5, fmt=PCT)
    ws.cell(r, 1).font = Font(name='Calibri', sz=10, color=INK)
    ws.cell(r, 2).font = Font(name='Calibri', sz=10, color=INK)
ws.cell(17, 2, 'Top ten together')
ws.cell(17, 3, '=SUM(C7:C16)'); ws.cell(17, 4, '=SUM(D7:D16)')
ws.cell(17, 5, "=IFERROR(D17/'P&L FY27'!$O$6,0)")
body(ws, 17, 3, 4, bold=True); body(ws, 17, 5, 5, fmt=PCT, bold=True)
rule(ws, 17, 2, 5)
ws.cell(18, 2, 'Everyone else')
ws.cell(18, 3, "='P&L FY27'!$D$6-C17"); ws.cell(18, 4, "='P&L FY27'!$O$6-D17")
ws.cell(18, 5, "=IFERROR(D18/'P&L FY27'!$O$6,0)")
body(ws, 18, 3, 4); body(ws, 18, 5, 5, fmt=PCT)
ws.cell(19, 2, 'Total income per the P&L')
ws.cell(19, 3, "='P&L FY27'!$D$6"); ws.cell(19, 4, "='P&L FY27'!$O$6")
body(ws, 19, 3, 4, bold=True); rule(ws, 19, 2, 5)
ws.cell(20, 2, 'of which income with no customer named')
ws.cell(20, 3, f'=SUMIFS({CA},{CJ},"",{CE},"Income",{CB},"="&{MON})')
ws.cell(20, 4, f'=SUMIFS({CA},{CJ},"",{CE},"Income",{CB},">="&{FY0},{CB},"<="&{MON})')
body(ws, 20, 3, 4)
note(ws.cell(20, 5), 'Journals, mostly work in progress')
ws.cell(22, 2, 'Check, the workings less the P&L, should be nil')
ws.cell(22, 4, f"=ROUND(SUM({RNG('C')})+$D$20-'P&L FY27'!$O$6,2)")
body(ws, 22, 4, 4, bold=True)
note(ws.cell(23, 2), 'If the check is not nil a customer has income in the GL that is not in the '
                     'list below. Add the name and it picks up on its own.')

banner(ws, W0 - 2, 1, 5, 'THE WORKINGS, EVERY CUSTOMER IN THE GL')
head(ws, W0 - 1, 1, ['Customer', 'This month', 'Year to date', 'Sort key', 'Rank'])
for i in range(W1 - W0 + 1):
    r = W0 + i
    if i < len(CUST):
        ws.cell(r, 1, CUST[i])
    ws.cell(r, 2, f'=SUMIFS({CA},{CJ},$A{r},{CE},"Income",{CB},"="&{MON})')
    ws.cell(r, 3, f'=SUMIFS({CA},{CJ},$A{r},{CE},"Income",{CB},">="&{FY0},{CB},"<="&{MON})')
    ws.cell(r, 4, f'=IF($A{r}="","",IF($C{r}<=0,"",$C{r}+ROW()/1000000))')
    ws.cell(r, 5, f'=IF($D{r}="","",RANK($C{r},{RNG("C")}))')
    body(ws, r, 2, 3)
    body(ws, r, 4, 5, fmt='General')
    ws.cell(r, 1).font = Font(name='Calibri', sz=9, color=INK)
    ws.cell(r, 4).font = Font(name='Calibri', sz=9, color=GREY)
    ws.cell(r, 5).font = Font(name='Calibri', sz=9, color=INK)
    ws.cell(r, 5).alignment = Alignment(horizontal='center')
ws.column_dimensions['A'].width = 46
for c in 'BCD':
    ws.column_dimensions[c].width = 15
ws.column_dimensions['E'].width = 13
ws.freeze_panes = 'A7'
print(f'Top 10 Clients: workings rows {W0}..{W1}, {len(CUST)} customers seeded')

# ========================================================== AR_Paste =======
ws = sheet('AR_Paste', 'BS_Paste')
ws.sheet_properties.tabColor = AQUA
for j, h in enumerate(AR['cols'], 1):
    ws.cell(1, j, h)
for i, row in enumerate(AR['rows'], 2):
    for j, v in enumerate(row, 1):
        ws.cell(i, j, v)
    ws.cell(i, 12, f'=SUM(F{i}:K{i})')
for i in range(2, len(AR['rows']) + 2):
    for j in range(6, 13):
        ws.cell(i, j).number_format = MONEY2
ws.column_dimensions['A'].width = 44
ws.column_dimensions['E'].width = 44
for c in 'BCD':
    ws.column_dimensions[c].width = 13
for j in range(6, 13):
    ws.column_dimensions[gl(j)].width = 13
AR_N = len(AR['rows']) + 1
print(f'AR_Paste: {AR_N-1} invoices seeded, rows 2..{AR_N}')

# ============================================================== DSO ========
ws = sheet('DSO', 'Utilisation')
A0, A1 = 2, 400
ARR = lambda c: f"AR_Paste!${c}${A0}:${c}${A1}"
title(ws['A1'], 'Days Sales Outstanding')
note(ws['A2'], 'Paste the Xero customer age analysis on AR_Paste, whole export, at A1. '
               'Everything below is worked out from it. The only thing you type is the target.')
ws['A3'] = '=TEXT(Setup!$B$5,"mmm yyyy")'
ws['A3'].font = Font(name='Calibri', sz=11, bold=True, color=TEAL)

banner(ws, 5, 1, 3, 'THE CALCULATION')
STEPS = [
 ('Trade receivables at month end', f'=SUM({ARR("L")})', MONEY, 'Total column of the age analysis'),
 ('Revenue, this month and the two before',
  "=SUMIFS('P&L FY27'!$B$6:$M$6,'P&L FY27'!$B$5:$M$5,\">=\"&EDATE(Setup!$B$5,-2),"
  "'P&L FY27'!$B$5:$M$5,\"<=\"&Setup!$B$5)", MONEY, 'Total income off the P&L'),
 ('Days in those three months', '=EOMONTH(Setup!$B$5,0)-EDATE(Setup!$B$5,-2)+1', '#,##0',
  'Calendar days, not working days'),
 ('Revenue a day', '=IF($B$8=0,0,$B$7/$B$8)', MONEY, 'Step 2 divided by step 3'),
 ('Days sales outstanding', '=IF($B$9=0,0,$B$6/$B$9)', DAYS, 'Receivables divided by revenue a day'),
]
r = 6
for lab, f, fmt, why in STEPS:
    ws.cell(r, 1, lab).font = Font(name='Calibri', sz=10, color=INK)
    ws.cell(r, 2, f); body(ws, r, 2, 2, fmt=fmt)
    note(ws.cell(r, 3), why)
    r += 1
ws.cell(11, 1, 'Target').font = Font(name='Calibri', sz=10, color=INK)
inp(ws, 'B11', 30, '#,##0')
note(ws.cell(11, 3), 'Type the target here. Everything below follows it.')
for lab, f, fmt, why in (
    ('Days above target', '=$B$10-$B$11', DAYS, 'How far off we are'),
    ('Cash locked up above target', '=$B$12*$B$9', MONEY,
     'Days above target times revenue a day. This is the cash a collection push releases.'),
    ('Receivables at the target', '=$B$11*$B$9', MONEY, 'What the book would be at target')):
    ws.cell(r + 1, 1, lab).font = Font(name='Calibri', sz=10, bold=True, color=INK)
    ws.cell(r + 1, 2, f); body(ws, r + 1, 2, 2, fmt=fmt, bold=True)
    note(ws.cell(r + 1, 3), why)
    r += 1
rule(ws, 13, 1, 2)

banner(ws, 17, 1, 4, 'AGEING')
head(ws, 18, 1, ['Bucket', 'A$', 'Share', 'Overdue'])
BUCK = [('Current', 'F', 'no'), ('Under 1 month', 'G', 'yes'), ('1 month', 'H', 'yes'),
        ('2 months', 'I', 'yes'), ('3 months', 'J', 'yes'), ('Over 3 months', 'K', 'yes')]
for i, (lab, col, od) in enumerate(BUCK):
    rr = 19 + i
    ws.cell(rr, 1, lab).font = Font(name='Calibri', sz=10, color=INK)
    ws.cell(rr, 2, f'=SUM({ARR(col)})')
    ws.cell(rr, 3, f'=IF($B$25=0,0,$B{rr}/$B$25)')
    ws.cell(rr, 4, 'Overdue' if od == 'yes' else '')
    body(ws, rr, 2, 2); body(ws, rr, 3, 3, fmt=PCT)
    ws.cell(rr, 4).font = Font(name='Calibri', sz=9, italic=True,
                               color=RED if od == 'yes' else INK)
ws.cell(25, 1, 'Total'); ws.cell(25, 2, '=SUM(B19:B24)'); ws.cell(25, 3, '=IF($B$25=0,0,1)')
body(ws, 25, 2, 2, bold=True); body(ws, 25, 3, 3, fmt=PCT, bold=True); rule(ws, 25, 1, 3)
ws.cell(26, 1, 'Overdue'); ws.cell(26, 2, '=SUM(B20:B24)')
ws.cell(26, 3, '=IF($B$25=0,0,$B$26/$B$25)')
body(ws, 26, 2, 2, bold=True); body(ws, 26, 3, 3, fmt=PCT, bold=True)
ws.cell(27, 1, 'Check against the balance sheet')
ws.cell(27, 2, "=ROUND($B$25-Balance_Sheet!$B$106,2)")
body(ws, 27, 2, 2, bold=True)
note(ws.cell(27, 3), 'Trade debtors on the balance sheet also carries other debtors, so a small '
                     'difference is normal. A large one means the paste is stale.')

CUSTS = sorted({x[0] for x in AR['rows']})
D0 = 32
D1 = D0 + max(len(CUSTS), 40) - 1
DR = lambda c: f'${c}${D0}:${c}${D1}'
banner(ws, 29, 1, 4, 'WHO OWES THE MOST')
head(ws, 30, 1, ['Rank', 'Customer', 'A$', 'Share'])
for n in range(1, 11):
    rr = 30 + n
    ws.cell(rr, 1, n).alignment = Alignment(horizontal='center')
    key = f'MATCH(LARGE({DR("H")},$A{rr}),{DR("H")},0)'
    ws.cell(rr, 2, f'=IFERROR(INDEX({DR("F")},{key}),"")')
    ws.cell(rr, 3, f'=IFERROR(INDEX({DR("G")},{key}),0)')
    ws.cell(rr, 4, f'=IF($B$25=0,0,$C{rr}/$B$25)')
    body(ws, rr, 3, 3); body(ws, rr, 4, 4, fmt=PCT)
    ws.cell(rr, 1).font = Font(name='Calibri', sz=10, color=INK)
    ws.cell(rr, 2).font = Font(name='Calibri', sz=10, color=INK)
banner(ws, D0 - 2, 6, 8, 'WORKINGS', AQUA, INK)
head(ws, D0 - 1, 6, ['Customer', 'Owed', 'Sort key'])
for i in range(D1 - D0 + 1):
    rr = D0 + i
    if i < len(CUSTS):
        ws.cell(rr, 6, CUSTS[i])
    ws.cell(rr, 7, f'=SUMIF({ARR("A")},$F{rr},{ARR("L")})')
    ws.cell(rr, 8, f'=IF($F{rr}="","",IF($G{rr}<=0,"",$G{rr}+ROW()/1000000))')
    body(ws, rr, 7, 7)
    ws.cell(rr, 6).font = Font(name='Calibri', sz=9, color=INK)
    ws.cell(rr, 8).font = Font(name='Calibri', sz=9, color=GREY)
ws.cell(D1 + 2, 6, 'Total'); ws.cell(D1 + 2, 7, f'=SUM({DR("G")})')
body(ws, D1 + 2, 7, 7, bold=True); rule(ws, D1 + 2, 6, 7)

T0 = D1 + 5
banner(ws, T0, 1, 14, 'DSO MONTH BY MONTH')
ws.cell(T0 + 1, 1, 'Month')
for j in range(12):
    c = ws.cell(T0 + 1, 2 + j, f"='P&L FY27'!{gl(2+j)}$5")
    c.number_format = 'mmm-yy'
    c.font = Font(name='Calibri', sz=10, bold=True, color=WHITE)
    c.fill = PatternFill('solid', fgColor=TEAL)
    c.alignment = Alignment(horizontal='center')
ws.cell(T0 + 1, 1).font = Font(name='Calibri', sz=10, bold=True, color=WHITE)
ws.cell(T0 + 1, 1).fill = PatternFill('solid', fgColor=TEAL)
ws.cell(T0 + 2, 1, 'Receivables at month end')
for j in range(12):
    cl = gl(2 + j)
    c = ws.cell(T0 + 2, 2 + j)
    c.number_format = MONEY
    c.fill = PatternFill('solid', fgColor=AQUA)
    c.font = Font(name='Calibri', sz=10, bold=True, color=TEAL)
ws.cell(T0 + 2, 4, '=$B$6')          # September, from the paste
ws.cell(T0 + 3, 1, 'Revenue, that month and the two before')
ws.cell(T0 + 4, 1, 'Days in those three months')
ws.cell(T0 + 5, 1, 'DSO')
for j in range(12):
    cl = gl(2 + j)
    ws.cell(T0 + 3, 2 + j,
            f"=SUMIFS('P&L FY27'!$B$6:$M$6,'P&L FY27'!$B$5:$M$5,\">=\"&EDATE({cl}${T0+1},-2),"
            f"'P&L FY27'!$B$5:$M$5,\"<=\"&{cl}${T0+1})")
    ws.cell(T0 + 4, 2 + j, f'=EOMONTH({cl}${T0+1},0)-EDATE({cl}${T0+1},-2)+1')
    ws.cell(T0 + 5, 2 + j,
            f'=IF(OR({cl}{T0+2}="",{cl}{T0+3}=0),"",{cl}{T0+2}/({cl}{T0+3}/{cl}{T0+4}))')
body(ws, T0 + 3, 2, 13); body(ws, T0 + 4, 2, 13, fmt='#,##0')
body(ws, T0 + 5, 2, 13, fmt=DAYS, bold=True)
for rr in (T0 + 2, T0 + 3, T0 + 4, T0 + 5):
    ws.cell(rr, 1).font = Font(name='Calibri', sz=10, color=INK)
note(ws.cell(T0 + 7, 1), 'Type the receivables balance into the row in teal at each month end. '
                         'September is already linked to the age analysis. The rest fill in as '
                         'the year goes, and the chart follows.')
ws.column_dimensions['A'].width = 38
for c in 'BCD':
    ws.column_dimensions[c].width = 16
ws.column_dimensions['F'].width = 44
ws.column_dimensions['G'].width = 14
ws.column_dimensions['H'].width = 14
print(f'DSO: calculation rows 6..14, ageing 19..27, debtors {D0}..{D1}, trend {T0}..{T0+5}')

json.dump({'ar_rows': AR_N, 'dso_trend': T0, 'top10_work': [W0, W1],
           'dso_debt': [D0, D1]}, open(f'{HERE}/data/newsheet_rows.json', 'w'))
wb.save(OUT)
print('saved', OUT)
