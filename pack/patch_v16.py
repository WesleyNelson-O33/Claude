"""Give the FY26 tabs a YTD column, and finish the Summary analysis block.

The YTD column is built the way P&L FY27 already builds its own. A money row
adds up the months to the chosen month. A margin or percentage row does not
add up: it works the ratio out again off the year to date money rows, which
is what P&L FY27 does and the only thing that gives a true year to date
margin.

Then the Summary. Her four headings read YTD Actual, YTD budget, YOY Actual,
YOY budget. Only the first had a formula, and that formula was a year on
year movement, not a year to date figure, so it sat under the wrong heading.
The last two were empty, which left Var $ and Var % reading off a blank cell
and showing minus one hundred per cent on every line.

Her note on U4 says "remove this YOY Budget", so that column goes. It held
the same figure the YOY budget column now holds.

  P  YTD Actual   this year to date, actual
  Q  YTD budget   this year to date, budget
  R  YOY Actual   this year to date less last year to date, actual
  S  YOY budget   this year to date less last year to date, budget
  U  Var $        R less S
  V  Var %        U over S

Same shape as the YOY block already on P&L FY27. Every cell takes its number
format from that row, so a margin row stays a percentage.
"""
import os, re, shutil, copy
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = f'{HERE}/CTS Financial Controller Pack - v14.xlsx'
OUT = f'{HERE}/CTS Financial Controller Pack - v16.xlsx'
R0, R1 = 6, 191

shutil.copy(SRC, OUT)
wb = openpyxl.load_workbook(OUT)

# ------------------------------------------- 1. YTD on the two FY26 tabs ----
NtoO = re.compile(r'\bN(\$?\d+)\b')
for name in ('Actual FY26', 'Budget FY26'):
    ws = wb[name]
    assert ws.cell(5, 14).value == 'FY total', (name, ws.cell(5, 14).value)
    assert ws.cell(5, 15).value is None, (name, 'column O is already used')
    ws.cell(5, 15)._style = copy.copy(ws.cell(5, 14)._style)
    ws.cell(5, 15).value = 'YTD'
    money = ratio = 0
    for r in range(R0, R1 + 1):
        n = ws.cell(r, 14).value
        if not isinstance(n, str) or not n.startswith('='):
            continue
        o = ws.cell(r, 15)
        if n.startswith('=SUM(B'):
            o.value = f"=SUMIF('{name}'!$B$5:$M$5,\"<=\"&Setup!$B$5,'{name}'!$B{r}:$M{r})"
            money += 1
        else:
            o.value = NtoO.sub(r'O\1', n)      # the ratio, worked out off the YTD column
            ratio += 1
        o._style = copy.copy(ws.cell(r, 14)._style)
    ws.column_dimensions['O'].width = ws.column_dimensions['N'].width
    print(f'  {name}: YTD written, {money} money rows, {ratio} ratio rows')

# ---------------------------------------- 2. the Summary analysis block ----
s = wb['Summary']
assert s['U4'].value.strip() == 'remove this YOY Budget', s['U4'].value
assert s['U5'].value.strip() == 'YOY Budget', s['U5'].value
note = s['V4'].value
s.delete_cols(21)
s['U4'] = note
assert s['U5'].value.strip() == 'Var $', s['U5'].value
assert s['W5'].value.strip() == 'Risk', s['W5'].value

n = 0
for r in range(R0, R1 + 1):
    if not isinstance(s.cell(r, 5).value, str):
        continue
    fmt = s.cell(r, 5).number_format            # money or percent, as that row is
    s.cell(r, 16, f'=IFERROR(E{r},"")')
    s.cell(r, 17, f'=IFERROR(G{r},"")')
    s.cell(r, 18, f"=IFERROR(E{r}-'Actual FY26'!O{r},\"\")")
    s.cell(r, 19, f"=IFERROR(G{r}-'Budget FY26'!O{r},\"\")")
    s.cell(r, 21, f'=IFERROR(R{r}-S{r},"")')
    s.cell(r, 22, f'=IFERROR(U{r}/ABS(S{r}),"")')
    for c in (16, 17, 18, 19, 21):
        s.cell(r, c)._style = copy.copy(s.cell(r, 5)._style)
        s.cell(r, c).number_format = fmt
    s.cell(r, 22).number_format = '0.0%;[Red](0.0%);\\-'
    n += 1
print('  Summary: analysis rows rebuilt:', n)
for w in ('Q', 'R', 'S'):
    s.column_dimensions[w].width = s.column_dimensions['P'].width

wb.save(OUT)
print('saved', OUT)
