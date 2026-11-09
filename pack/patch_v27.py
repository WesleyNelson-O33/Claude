"""Rolling 3 Months: the three months ending at whatever the dropdown says.

Pick October and it shows August, September, October. Pick November and it
rolls to September, October, November.

Every line of the P&L, laid out the same way, with the same + and - grouping.
Beside the three months it carries the three month total, the three months
before that, and the change between them.

Where the window reaches back past the start of the financial year it picks
the month up off the Actual FY26 tab, which is where last year's months live.
"""
import os, json, copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.properties import Outline
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = f'{HERE}/CTS Financial Controller Pack - v20.xlsx'
B = json.load(open(f'{HERE}/data/brand.json'))
INK, WHITE, GREY, GREYL, AQUA = (B['charcoal'], B['white'], B['grey'],
                                 B['grey_lt'], B['bright_teal'])
PCT = '0.0%;[Red](0.0%);\\-'
VAR = '0.0%;[Red](0.0%);\\-'
FIRST, LAST = 6, 181                      # the P&L line block
REV = 6                                   # Total Income
RATIO = {174: 173, 177: 176, 179: 178, 181: 180}   # a % row and its money row

wb = openpyxl.load_workbook(OUT)
PL = wb['P&L FY27']
if 'Rolling 3 Months' in wb.sheetnames:
    del wb['Rolling 3 Months']
ws = wb.create_sheet('Rolling 3 Months', wb.sheetnames.index('Summary') + 1)
ws.sheet_properties.tabColor = INK
ws.sheet_properties.outlinePr = Outline(applyStyles=False, summaryBelow=False,
                                        summaryRight=False)

ws['A1'] = 'Rolling three months'
ws['A1'].font = Font(name='Calibri', size=16, bold=True, color=INK)
ws['A2'] = ('=TEXT(EDATE(Setup!$B$5,-2),"mmm yyyy")&" to "&'
            'TEXT(Setup!$B$5,"mmm yyyy")')
ws['A2'].font = Font(name='Calibri', size=10, bold=True, color=INK)

HDR = 5
hair = Side(style='medium', color=AQUA)
for i, lab in enumerate(['Line', None, None, None, '3 months',
                         'Prior 3 months', 'Change', 'Change %'], start=1):
    c = ws.cell(HDR, i, lab)
    c.font = Font(name='Calibri', size=11, bold=True, color=WHITE)
    c.fill = PatternFill('solid', fgColor=INK)
    c.alignment = Alignment(horizontal='right' if i > 1 else 'left')
    c.border = Border(bottom=hair)
for j, off in enumerate((-2, -1, 0)):
    c = ws.cell(HDR, 2 + j, f'=EDATE(Setup!$B$5,{off})')
    c.number_format = 'mmm\\-yy'

# a month, wherever in the pack it lives
FY, F26 = "'P&L FY27'", "'Actual FY26'"
def val(row, mexpr):
    return (f'IFERROR(IF({mexpr}>=DATE(Setup!$B$7-1,Setup!$B$6,1),'
            f'INDEX({FY}!$B{row}:$M{row},MATCH({mexpr},{FY}!$B$5:$M$5,0)),'
            f'INDEX({F26}!$B{row}:$M{row},MATCH(EDATE({mexpr},12),{F26}!$B$5:$M$5,0))),0)')

def three(row, back):
    """The three months ending `back` months before the selected one."""
    return '+'.join(val(row, f'EDATE(Setup!$B$5,{-(back + k)})') for k in (2, 1, 0))

n = 0
for r in range(FIRST, LAST + 1):
    src = PL.cell(r, 1)
    ws.cell(r, 1, src.value)
    ws.cell(r, 1)._style = copy.copy(src._style)
    d = PL.row_dimensions.get(r)
    if d is not None:
        ws.row_dimensions[r].outlineLevel = d.outlineLevel
        ws.row_dimensions[r].hidden = d.hidden
    style = copy.copy(PL.cell(r, 2)._style)
    if r in RATIO:
        m = RATIO[r]
        for j, off in enumerate((-2, -1, 0)):
            ws.cell(r, 2 + j, f'=IF($B{REV}+0=0,"",{gl(2+j)}{m}/{gl(2+j)}{REV})')
        ws.cell(r, 5, f'=IF($E{REV}=0,"",$E{m}/$E{REV})')
        ws.cell(r, 6, f'=IF($F{REV}=0,"",$F{m}/$F{REV})')
        for j in range(3):
            ws.cell(r, 2 + j, f'=IF({gl(2+j)}{REV}=0,"",{gl(2+j)}{m}/{gl(2+j)}{REV})')
        fmt = PCT
    else:
        for j, off in enumerate((-2, -1, 0)):
            ws.cell(r, 2 + j, '=' + val(r, f'EDATE(Setup!$B$5,{off})'))
        ws.cell(r, 5, f'=SUM($B{r}:$D{r})')
        ws.cell(r, 6, '=' + three(r, 3))
        fmt = PL.cell(r, 2).number_format
    ws.cell(r, 7, f'=IF(OR($E{r}="",$F{r}=""),"",$E{r}-$F{r})')
    ws.cell(r, 8, f'=IF(OR($F{r}="",$F{r}=0),"",$G{r}/ABS($F{r}))')
    for c in range(2, 8):
        ws.cell(r, c)._style = copy.copy(style)
        ws.cell(r, c).number_format = fmt
    ws.cell(r, 8)._style = copy.copy(style)
    ws.cell(r, 8).number_format = VAR
    n += 1

ws.column_dimensions['A'].width = 38
for c in 'BCDEFGH':
    ws.column_dimensions[c].width = 15
ws.freeze_panes = 'B6'
print(f'Rolling 3 Months built: rows {FIRST} to {LAST}, {n} lines')
wb.save(OUT)
print('saved')
