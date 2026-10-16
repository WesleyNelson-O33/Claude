"""Two new sheets.

Top 10 YoY      the year on year block, lifted off the Top 10 Clients sheet
                onto its own tab, with its workings and its chart
Overheads MoM   every overhead account, month by month across the year, with
                what moved on last month
"""
import os, json, copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.properties import Outline
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = f'{HERE}/CTS Financial Controller Pack - v20.xlsx'
B = json.load(open(f'{HERE}/data/brand.json'))
FY26 = json.load(open(f'{HERE}/data/fy26_customers.json'))
ACC = json.load(open(f'{HERE}/data/overhead_accounts.json'))

INK, WHITE, GREY, GREYL = B['charcoal'], B['white'], B['grey'], B['grey_lt']
AQUA, NAVY, BLUE, PURP, TEAL = (B['bright_teal'], B['deep_blue'], B['primary_blue'],
                                B['purple'], B['dark_teal'])
MONEY = '\\$#,##0;[Red]"($"#,##0\\);\\-'
COST  = '[Red]"($"#,##0\\);\\$#,##0;\\-'
PCT   = '0.0%;[Red](0.0%);\\-'
MOVE  = '[Red]\\+#,##0;\\(#,##0\\);\\-'        # spend up is red, spend down in brackets
thin  = Side(style='thin', color=GREY)
hair  = Side(style='medium', color=AQUA)

wb = openpyxl.load_workbook(OUT)

def title(ws, text):
    ws['A1'] = text
    ws['A1'].font = Font(name='Calibri', size=16, bold=True, color=INK)

def head(ws, r, labels, widths=None):
    for i, t in enumerate(labels, start=1):
        c = ws.cell(r, i, t)
        c.font = Font(name='Calibri', size=10, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=INK)
        c.alignment = Alignment(horizontal='right' if i > 2 else 'left', wrap_text=True)
        c.border = Border(bottom=hair)

def band(ws, r, ncol, fill=GREY, bold=True):
    for c in range(1, ncol + 1):
        cell = ws.cell(r, c)
        cell.fill = PatternFill('solid', fgColor=fill)
        cell.font = Font(name='Calibri', size=11, bold=bold, color=INK)
        cell.border = Border(bottom=thin)
        if c > 1:
            cell.alignment = Alignment(horizontal='right')

def plain(ws, r, ncol, bold=False):
    for c in range(1, ncol + 1):
        cell = ws.cell(r, c)
        cell.font = Font(name='Calibri', size=11, bold=bold, color=INK)
        cell.border = Border(bottom=thin)
        if c > 1:
            cell.alignment = Alignment(horizontal='right')

# ======================================================= 1. Top 10 YoY =====
for nm in ('Top 10 YoY',):
    if nm in wb.sheetnames:
        del wb[nm]
ws = wb.create_sheet('Top 10 YoY', wb.sheetnames.index('Top 10 Clients') + 1)
ws.sheet_properties.tabColor = INK
T10 = "'Top 10 Clients'"
title(ws, 'Top 10 Clients, year on year')
ws['A2'] = '=TEXT(Setup!$B$5,"mmm yyyy")'
ws['A2'].font = Font(name='Calibri', size=10, bold=True, color=INK)

NF = len(FY26)
DUMP = 34
FR, FV = f'$A${DUMP+1}:$A${DUMP+NF}', f'$B${DUMP+1}:$B${DUMP+NF}'

head(ws, 4, ['Rank', 'Customer', 'FY27 YTD', 'FY26 YTD', 'Change', 'Change %',
             'FY26 rank', 'Movement'])
for i in range(10):
    r, src = 5 + i, 7 + i
    ws.cell(r, 1, i + 1)
    ws.cell(r, 2, f'=IF({T10}!$B{src}="","",{T10}!$B{src})')
    ws.cell(r, 3, f'=IF($B{r}="","",{T10}!$D{src})')
    ws.cell(r, 4, f'=IF($B{r}="","",IFERROR(SUMIF({FR},$B{r},{FV}),0))')
    ws.cell(r, 5, f'=IF($B{r}="","",$C{r}-$D{r})')
    ws.cell(r, 6, f'=IF($B{r}="","",IF($D{r}=0,"",$E{r}/$D{r}))')
    ws.cell(r, 7, f'=IF($B{r}="","",IF($D{r}=0,"",COUNTIF({FV},">"&$D{r})+1))')
    ws.cell(r, 8, f'=IF($B{r}="","",IF($D{r}=0,"new this year",IF($G{r}=$A{r},"same",'
                  f'IF($G{r}>$A{r},"up "&$G{r}-$A{r},"down "&$A{r}-$G{r}))))')
    plain(ws, r, 8)
    for c in (3, 4, 5):
        ws.cell(r, c).number_format = MONEY
    ws.cell(r, 6).number_format = PCT

ws.cell(15, 2, 'Top ten together')
ws.cell(15, 3, '=SUM($C$5:$C$14)')
ws.cell(15, 4, '=SUM($D$5:$D$14)')
ws.cell(15, 5, '=$C15-$D15')
ws.cell(15, 6, '=IF($D15=0,"",$E15/$D15)')
ws.cell(15, 8, '=COUNTIF($H$5:$H$14,"new this year")&" new, "'
               '&COUNTIF($H$5:$H$14,"same")&" held rank"')
band(ws, 15, 8, GREYL)
for c in (3, 4, 5):
    ws.cell(15, c).number_format = MONEY
ws.cell(15, 6).number_format = PCT

ws.cell(17, 1, 'TOP TEN LAST YEAR THAT IS NOT TOP TEN NOW')
band(ws, 17, 8, INK)
for c in range(1, 9):
    ws.cell(17, c).font = Font(name='Calibri', size=11, bold=True, color=WHITE)
head(ws, 18, ['FY26 rank', 'Customer', 'FY27 YTD', 'FY26 YTD', 'Change', 'Change %',
              'Still a client', ''])
drop = sorted(FY26.items(), key=lambda x: -x[1])[:10]
for i, (n, v) in enumerate(drop, start=1):
    r = 18 + i
    ws.cell(r, 1, i)
    ws.cell(r, 2, n)
    ws.cell(r, 3, f'=IFERROR(INDEX({T10}!$C$24:$C$103,MATCH($B{r},{T10}!$A$24:$A$103,0)),0)')
    ws.cell(r, 4, f'=IFERROR(SUMIF({FR},$B{r},{FV}),0)')
    ws.cell(r, 5, f'=$C{r}-$D{r}')
    ws.cell(r, 6, f'=IF($D{r}=0,"",$E{r}/$D{r})')
    ws.cell(r, 7, f'=IF(COUNTIF($B$5:$B$14,$B{r})>0,"still top ten",'
                  f'IF($C{r}>0,"yes, outside the top ten","no revenue this year"))')
    plain(ws, r, 8)
    for c in (3, 4, 5):
        ws.cell(r, c).number_format = MONEY
    ws.cell(r, 6).number_format = PCT

ws.cell(DUMP - 1, 1, 'THE FY26 WORKINGS, EVERY CUSTOMER INVOICED JUL TO SEP 2025, EX GST')
band(ws, DUMP - 1, 8, INK)
for c in range(1, 9):
    ws.cell(DUMP - 1, c).font = Font(name='Calibri', size=11, bold=True, color=WHITE)
head(ws, DUMP, ['Customer', 'FY26 YTD', '', '', '', '', '', ''])
for i, (n, v) in enumerate(sorted(FY26.items(), key=lambda x: -x[1])):
    ws.cell(DUMP + 1 + i, 1, n)
    ws.cell(DUMP + 1 + i, 2, v).number_format = MONEY
ws.column_dimensions['A'].width = 46
for c, w in (('B', 40), ('C', 14), ('D', 14), ('E', 13), ('F', 11), ('G', 11), ('H', 22)):
    ws.column_dimensions[c].width = w
ws.freeze_panes = 'A5'
print('Top 10 YoY built')

# take the block, and its chart, off the Top 10 Clients sheet
old = wb['Top 10 Clients']
old._charts = [c for c in old._charts if c.anchor._from.row < 100]
old.delete_rows(105, old.max_row - 104)
print('Top 10 Clients trimmed back to', old.max_row, 'rows,', len(old._charts), 'chart')

# ==================================================== 2. Overheads MoM =====
if 'Overheads MoM' in wb.sheetnames:
    del wb['Overheads MoM']
ws = wb.create_sheet('Overheads MoM', wb.sheetnames.index('Overhead Split') + 1)
ws.sheet_properties.tabColor = INK
ws.sheet_properties.outlinePr = Outline(applyStyles=False, summaryBelow=False,
                                        summaryRight=False)
title(ws, 'Overheads, month by month')
ws['A2'] = '=TEXT(Setup!$B$5,"mmm yyyy")'
ws['A2'].font = Font(name='Calibri', size=10, bold=True, color=INK)

NC = 17                                   # A plus 12 months plus 4
HR = 4
ws.cell(HR, 1, 'Line')
for i in range(12):
    ws.cell(HR, 2 + i, f"='P&L FY27'!{gl(2+i)}$5").number_format = 'mmm\\-yy'
ws.cell(HR, 14, 'Year to date')
ws.cell(HR, 15, 'Monthly average')
ws.cell(HR, 16, 'Spend up / (down) on last month')
ws.cell(HR, 17, 'On last month %')
head(ws, HR, [ws.cell(HR, c).value for c in range(1, NC + 1)])
for i in range(12):
    ws.cell(HR, 2 + i, f"='P&L FY27'!{gl(2+i)}$5").number_format = 'mmm\\-yy'
ws.row_dimensions[HR].height = 30

CA, CB, CE, CF, CI = ('Cleanup!$A$5:$A$20004', 'Cleanup!$B$5:$B$20004',
                      'Cleanup!$E$5:$E$20004', 'Cleanup!$F$5:$F$20004',
                      'Cleanup!$I$5:$I$20004')
MTH = f'MATCH(Setup!$B$5,$B${HR}:$M${HR},0)'

def tail(r):
    ws.cell(r, 14, f'=SUMIF($B${HR}:$M${HR},"<="&Setup!$B$5,$B{r}:$M{r})')
    ws.cell(r, 15, f'=IFERROR($N{r}/Setup!$B$8,"")')
    ws.cell(r, 16, f'=IF(Setup!$B$8=1,"",INDEX($B{r}:$M{r},{MTH}-1)'
                   f'-INDEX($B{r}:$M{r},{MTH}))')
    ws.cell(r, 17, f'=IF(OR(Setup!$B$8=1,INDEX($B{r}:$M{r},{MTH}-1)=0),"",'
                   f'$P{r}/ABS(INDEX($B{r}:$M{r},{MTH}-1)))')
    for c in (14, 15):
        ws.cell(r, c).number_format = COST
    ws.cell(r, 16).number_format = MOVE
    ws.cell(r, 17).number_format = PCT

TOT = HR + 1
ws.cell(TOT, 1, 'Total overheads')
SUBROWS = []
r = TOT + 2
for sub, accs in ACC.items():
    SUBROWS.append(r)
    sub_row = r
    ws.cell(r, 1, sub)
    for i in range(12):
        ws.cell(r, 2 + i, f'=SUMIFS({CA},{CF},TRIM($A{r}),{CE},"Expenses",'
                          f'{CB},{gl(2+i)}${HR})').number_format = COST
    tail(r)
    band(ws, r, NC, GREYL)
    r += 1
    seen = set()
    for code, name in accs:
        if name in seen:
            continue
        seen.add(name)
        ws.cell(r, 1, '    ' + name)
        # one account name is used under two subcategories, so pin the leaf
        # to its own heading or it gets counted in both
        for i in range(12):
            ws.cell(r, 2 + i, f'=SUMIFS({CA},{CI},TRIM($A{r}),{CF},TRIM($A${sub_row}),'
                              f'{CE},"Expenses",{CB},{gl(2+i)}${HR})').number_format = COST
        tail(r)
        plain(ws, r, NC)
        ws.row_dimensions[r].outlineLevel = 1
        ws.row_dimensions[r].hidden = True
        r += 1

for i in range(12):
    ws.cell(TOT, 2 + i, '=' + '+'.join(f'{gl(2+i)}{x}' for x in SUBROWS)).number_format = COST
tail(TOT)
band(ws, TOT, NC, GREY)

CHK = r + 1
# Only the months that have actually happened. From next month on the P&L
# carries budget, and this sheet carries the GL, so they are not comparable.
ws.cell(CHK, 1, 'Check against the P&L overhead line, months to date')
for i in range(12):
    ws.cell(CHK, 2 + i,
            f'=IF({gl(2+i)}${HR}>Setup!$B$5,"",'
            f"ROUND({gl(2+i)}{TOT}-'P&L FY27'!{gl(2+i)}$164,2))").number_format = COST
plain(ws, CHK, 13, bold=True)

ws.column_dimensions['A'].width = 38
for i in range(12):
    ws.column_dimensions[gl(2 + i)].width = 12
for c, w in (('N', 14), ('O', 14), ('P', 16), ('Q', 13)):
    ws.column_dimensions[c].width = w
ws.freeze_panes = 'B5'
print(f'Overheads MoM built: total row {TOT}, {len(SUBROWS)} subcategories, '
      f'detail to row {r-1}, check row {CHK}')

wb.save(OUT)
print('saved')
