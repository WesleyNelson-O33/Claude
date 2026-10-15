"""Put the overhead allocation on the P&L, month by month.

Two blocks go in under the margin block on P&L FY27. The first shows what
each department is charged for admin overhead in each month, straight off
the Overhead Split matrix. The second shows what each department is left
with after it: its own income, less its own direct costs, less its share of
the overhead.

Nothing above row 191 is touched, so every figure already on the P&L stays
exactly as it was.
"""
import os, json, shutil, copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = f'{HERE}/CTS Financial Controller Pack - v17.xlsx'
OUT = f'{HERE}/CTS Financial Controller Pack - v18.xlsx'
meta = json.load(open(f'{HERE}/data/overhead_rows.json'))
A0 = meta['alloc_first_row']
W0 = meta['own_first_row']

TEAL, GREY_LT, MID, INK, WHITE = '005358', 'F4F4F4', '5F6670', '1A1A1A', 'FFFFFF'
MONEY = '\\$#,##0;[Red]"($"#,##0\\);\\-'
PCT = '0.0%;[Red](0.0%);\\-'

# department, income row, direct cost row, row on Overhead Split
DEPT = [('Support',     41, 130, A0 + 0, W0 + 0),
        ('Events',       8,  61, A0 + 1, W0 + 1),
        ('Video',       16,  78, A0 + 2, W0 + 2),
        ('Integration', 25,  96, A0 + 3, W0 + 3),
        ('Consulting',  33, 113, A0 + 4, W0 + 4)]

shutil.copy(SRC, OUT)
wb = openpyxl.load_workbook(OUT)
ws = wb['P&L FY27']

for name, inc, cos, _, _w in DEPT:
    assert str(ws.cell(inc, 1).value).strip() == name, (name, ws.cell(inc, 1).value)
    assert str(ws.cell(cos, 1).value).strip() == name, (name, ws.cell(cos, 1).value)
assert str(ws.cell(183, 1).value).strip() == 'Profit Margin Per Department'
assert ws.cell(192, 1).value is None, 'row 192 is not free'

BANNER = copy.copy(ws.cell(183, 1)._style)        # the margin block banner
LEAF = copy.copy(ws.cell(185, 1)._style)          # an indented leaf row
LEAFV = copy.copy(ws.cell(185, 2)._style)

def banner(r, text):
    for c in range(1, 16):
        ws.cell(r, c)._style = copy.copy(ws.cell(183, c)._style)
    ws.cell(r, 1, text)

def block(r0, text, value_of, fmt):
    banner(r0, text)
    for i, (name, inc, cos, orow, wrow) in enumerate(DEPT):
        r = r0 + 1 + i
        ws.cell(r, 1, '        ' + name)._style = LEAF
        for j in range(12):
            c = ws.cell(r, 2 + j)
            c.value = value_of(gl(2 + j), gl(3 + j), inc, cos, orow, wrow, r)
            c._style = copy.copy(ws.cell(185, 2 + j)._style)
            c.number_format = fmt
        n = ws.cell(r, 14, f'=SUM(B{r}:M{r})')
        n._style = copy.copy(ws.cell(185, 14)._style)
        n.number_format = fmt
        o = ws.cell(r, 15, f"=SUMIF('P&L FY27'!$B$5:$M$5,\"<=\"&Setup!$B$5,'P&L FY27'!$B{r}:$M{r})")
        o._style = copy.copy(ws.cell(185, 15)._style)
        o.number_format = fmt
    return r0 + 5

R = 193
end = block(R, "Each department's own overhead",
            lambda pc, oc, inc, cos, orow, wrow, r: f"='Overhead Split'!{oc}{wrow}", MONEY)

RA = end + 2
endA = block(RA, 'Admin overhead charged across',
             lambda pc, oc, inc, cos, orow, wrow, r: f"='Overhead Split'!{oc}{orow}", MONEY)
end = endA

R2 = end + 2
end2 = block(R2, 'Result after overhead',
             lambda pc, oc, inc, cos, orow, wrow, r:
                 f'={pc}{inc}+{pc}{cos}+{pc}{r - (R2 - R)}+{pc}{r - (R2 - RA)}', MONEY)

R3 = end2 + 2
banner(R3, 'Margin after overhead')
for i, (name, inc, cos, orow, wrow) in enumerate(DEPT):
    r = R3 + 1 + i
    res = r - (R3 - R2)
    ws.cell(r, 1, '        ' + name)._style = LEAF
    for j in range(12):
        col = gl(2 + j)
        c = ws.cell(r, 2 + j, f'=IF({col}{inc}=0,"",{col}{res}/{col}{inc})')
        c._style = copy.copy(ws.cell(185, 2 + j)._style)
        c.number_format = PCT
    for cc, src in ((14, 14), (15, 15)):
        col = gl(cc)
        c = ws.cell(r, cc, f'=IF({col}{inc}=0,"",{col}{res}/{col}{inc})')
        c._style = copy.copy(ws.cell(185, src)._style)
        c.number_format = PCT

ws.cell(R3 + 7, 1, 'The split comes off the Overhead Split tab. Set the headcount and '
                   'the two typed keys there once a year.')
ws.cell(R3 + 7, 1).font = Font(name='Calibri', sz=9, italic=True, color=MID)
print(f'P&L FY27: overhead charged {R}..{end}, result {R2}..{end2}, margin {R3}..{R3+5}')

wb.save(OUT)
print('saved', OUT)
