"""Top ten year on year, on the Top 10 Clients sheet.

Who was top last year, who is top now, and who moved. FY27 is earned revenue
off the Cleanup tab, the same basis as the block above. FY26 is the 145
invoices Xero issued 1 Jul to 30 Sep 2025, ex GST, dumped below so the
workings are visible.
"""
import os, json, copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = f'{HERE}/CTS Financial Controller Pack - v20.xlsx'
B = json.load(open(f'{HERE}/data/brand.json'))
FY26 = json.load(open(f'{HERE}/data/fy26_customers.json'))

TEAL, AQUA, BLUE, INK, WHITE = (B['dark_teal'], B['bright_teal'], B['primary_blue'],
                                B['charcoal'], B['white'])
RED, GREY, GREYL = B['accent_red'], B['grey'], B['grey_lt']
MONEY = '\\$#,##0;[Red]"($"#,##0\\);\\-'
PCT = '0.0%;[Red](0.0%);\\-'

wb = openpyxl.load_workbook(OUT)
ws = wb['Top 10 Clients']

thin = Side(style='thin', color=GREY)
def banner(r, text):
    c = ws.cell(r, 1, text)
    c.font = Font(name='Calibri', size=11, bold=True, color=WHITE)
    for cc in range(1, 9):
        ws.cell(r, cc).fill = PatternFill('solid', fgColor=TEAL)
        ws.cell(r, cc).font = Font(name='Calibri', size=11, bold=True, color=WHITE)

def header(r, labels):
    for i, t in enumerate(labels, start=1):
        c = ws.cell(r, i, t)
        c.font = Font(name='Calibri', size=10, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=INK)
        c.alignment = Alignment(horizontal='right' if i > 2 else 'left')
        c.border = Border(bottom=thin)

R0 = 106                      # the year on year block starts here
DUMP = R0 + 30                # the FY26 workings below it
NF = len(FY26)
FR = f'$A${DUMP + 1}:$A${DUMP + NF}'     # FY26 names
FV = f'$B${DUMP + 1}:$B${DUMP + NF}'     # FY26 revenue

banner(R0, 'TOP TEN YEAR ON YEAR')
ws.cell(R0 + 1, 1, 'FY27 is this year to date off the GL. FY26 is the same three months '
                   'last year, from the Xero invoices dumped below, ex GST.')
ws.cell(R0 + 1, 1).font = Font(name='Calibri', size=9, italic=True, color=INK)
header(R0 + 2, ['Rank', 'Customer', 'FY27 YTD', 'FY26 YTD', 'Change', 'Change %',
                'FY26 rank', 'Movement'])

for i in range(10):
    r = R0 + 3 + i
    src = 7 + i                                   # the top ten block above
    ws.cell(r, 1, i + 1)
    ws.cell(r, 2, f'=IF($B{src}="","",$B{src})')
    ws.cell(r, 3, f'=IF($B{r}="","",$D{src})')
    ws.cell(r, 4, f'=IF($B{r}="","",IFERROR(SUMIF({FR},$B{r},{FV}),0))')
    ws.cell(r, 5, f'=IF($B{r}="","",$C{r}-$D{r})')
    ws.cell(r, 6, f'=IF($B{r}="","",IF($D{r}=0,"",$E{r}/$D{r}))')
    ws.cell(r, 7, f'=IF($B{r}="","",IF($D{r}=0,"",COUNTIF({FV},">"&$D{r})+1))')
    ws.cell(r, 8, f'=IF($B{r}="","",IF($D{r}=0,"new this year",'
                  f'IF($G{r}=$A{r},"same",IF($G{r}>$A{r},"up "&$G{r}-$A{r},'
                  f'"down "&$A{r}-$G{r}))))')

T = R0 + 13
ws.cell(T, 2, 'Top ten together')
ws.cell(T, 3, f'=SUM($C{R0+3}:$C{R0+12})')
ws.cell(T, 4, f'=SUM($D{R0+3}:$D{R0+12})')
ws.cell(T, 5, f'=$C{T}-$D{T}')
ws.cell(T, 6, f'=IF($D{T}=0,"",$E{T}/$D{T})')
ws.cell(T, 8, f'=COUNTIF($H{R0+3}:$H{R0+12},"new this year")&" new, "'
              f'&COUNTIF($H{R0+3}:$H{R0+12},"same")&" held rank"')

# who was top ten last year and is not now --------------------------------
S = R0 + 15
banner(S, 'TOP TEN LAST YEAR THAT IS NOT TOP TEN NOW')
header(S + 1, ['FY26 rank', 'Customer', 'FY27 YTD', 'FY26 YTD', 'Change', 'Change %',
               'Still a client', ''])
drop = [(n, v) for n, v in sorted(FY26.items(), key=lambda x: -x[1])[:10]]
row = S + 2
for i, (n, v) in enumerate(drop, start=1):
    ws.cell(row, 1, i)
    ws.cell(row, 2, n)
    ws.cell(row, 3, f'=IFERROR(INDEX($C$24:$C$103,MATCH($B{row},$A$24:$A$103,0)),0)')
    ws.cell(row, 4, f'=IFERROR(SUMIF({FR},$B{row},{FV}),0)')
    ws.cell(row, 5, f'=$C{row}-$D{row}')
    ws.cell(row, 6, f'=IF($D{row}=0,"",$E{row}/$D{row})')
    ws.cell(row, 7, f'=IF(COUNTIF($B${R0+3}:$B${R0+12},$B{row})>0,"still top ten",'
                    f'IF($C{row}>0,"yes, outside the top ten","no revenue this year"))')
    row += 1

# the FY26 workings ---------------------------------------------------------
banner(DUMP - 1, 'THE FY26 WORKINGS, EVERY CUSTOMER INVOICED JUL TO SEP 2025')
header(DUMP, ['Customer', 'FY26 YTD, ex GST', '', '', '', '', '', ''])
for i, (n, v) in enumerate(sorted(FY26.items(), key=lambda x: -x[1])):
    ws.cell(DUMP + 1 + i, 1, n)
    ws.cell(DUMP + 1 + i, 2, v)

# formats -------------------------------------------------------------------
for r in list(range(R0 + 3, R0 + 14)) + list(range(S + 2, S + 12)):
    for c in (3, 4, 5):
        ws.cell(r, c).number_format = MONEY
    ws.cell(r, 6).number_format = PCT
    for c in range(1, 9):
        ws.cell(r, c).border = Border(bottom=thin)
        if c >= 3:
            ws.cell(r, c).alignment = Alignment(horizontal='right')
for r in (T,):
    for c in range(1, 9):
        ws.cell(r, c).font = Font(name='Calibri', size=11, bold=True, color=INK)
        ws.cell(r, c).fill = PatternFill('solid', fgColor=GREYL)
for i in range(NF):
    ws.cell(DUMP + 1 + i, 2).number_format = MONEY
for col, w in (('F', 12), ('G', 11), ('H', 20), ('C', 14), ('D', 14), ('E', 13)):
    ws.column_dimensions[col].width = w

wb.save(OUT)
print(f'Top ten year on year written: rows {R0} to {row}, FY26 dump {DUMP+1} to {DUMP+NF}')
