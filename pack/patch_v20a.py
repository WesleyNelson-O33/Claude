"""Four new sheets, and the CTS palette she supplied applied across the book.

  P&L by Dept     every P&L line, each department down a column, month and YTD
  Top 10 Clients  the ten biggest customers, with the workings they come from
  AR_Paste        a bare dump for the Xero customer age analysis
  DSO             days sales outstanding, worked out step by step off that dump

Everything on them is a formula. Nothing is a typed figure except the dump
itself and the four things on Setup she already types.
"""
import os, json, shutil, copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = f'{HERE}/CTS Financial Controller Pack - v19.xlsx'
OUT = f'{HERE}/CTS Financial Controller Pack - v20.xlsx'
B = json.load(open(f'{HERE}/data/brand.json'))
NS = json.load(open(f'{HERE}/data/newsheets.json'))
AR = json.load(open(f'{HERE}/data/ar_sep26.json'))

TEAL, AQUA, BLUE, NAVY = B['dark_teal'], B['bright_teal'], B['primary_blue'], B['deep_blue']
RED, DRED, PURP = B['accent_red'], B['deep_red'], B['purple']
INK, WHITE, GREY, GREYL = B['charcoal'], B['white'], B['grey'], B['grey_lt']
MONEY = '\\$#,##0;[Red]"($"#,##0\\);\\-'
MONEY2 = '\\$#,##0.00;[Red]"($"#,##0.00\\);\\-'
PCT = '0.0%;[Red](0.0%);\\-'
NUM1 = '#,##0.0;[Red](#,##0.0);\\-'

CA, CB, CD = 'Cleanup!$A$5:$A$20004', 'Cleanup!$B$5:$B$20004', 'Cleanup!$D$5:$D$20004'
CE, CF, CJ = 'Cleanup!$E$5:$E$20004', 'Cleanup!$F$5:$F$20004', 'Cleanup!$J$5:$J$20004'
FY0 = 'DATE(Setup!$B$7-1,Setup!$B$6,1)'          # first day of the financial year
MON = 'Setup!$B$5'

DEPTS = [('Support', 'ONSITE'), ('Events', 'PRODUCTION'), ('Video', 'VIDEO'),
         ('Integration', 'INTEGRATION'), ('Consulting', 'CONSULTING'),
         ('General Admin', 'ADMIN')]

shutil.copy(SRC, OUT)
wb = openpyxl.load_workbook(OUT)

# ------------------------------------------------- the contact, on Cleanup --
cu = wb['Cleanup']
cu['J4'] = 'Contact'
cu['J4']._style = copy.copy(cu['I4']._style)
for r in range(5, 20005):
    g = r - 4
    cu.cell(r, 10, f'=IF(N(\'P&L Analysis\'!$D{g})=0,"",TRIM(\'P&L Analysis\'!$E{g}))')
    cu.cell(r, 10)._style = copy.copy(cu.cell(r, 9)._style)
cu.column_dimensions['J'].width = 34
print('Cleanup: contact column written')

# --------------------------------------------------------------- helpers ---
def sheet(name, after):
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name, wb.sheetnames.index(after) + 1)
    ws.sheet_properties.tabColor = TEAL
    ws.sheet_view.showGridLines = False
    return ws

def title(c, t, sz=14):
    c.value = t
    c.font = Font(name='Calibri', sz=sz, bold=True, color=TEAL)

def note(c, t):
    c.value = t
    c.font = Font(name='Calibri', sz=9, italic=True, color=INK)

def head(ws, r, c0, labels, h=22):
    for i, t in enumerate(labels):
        c = ws.cell(r, c0 + i, t)
        c.font = Font(name='Calibri', sz=10, bold=True, color=WHITE)
        c.fill = PatternFill('solid', fgColor=TEAL)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[r].height = h

def band(ws, r, c0, c1, colour=AQUA, ink=INK):
    for c in range(c0, c1 + 1):
        x = ws.cell(r, c)
        x.fill = PatternFill('solid', fgColor=colour)
        x.font = Font(name='Calibri', sz=10, bold=True, color=ink)

def tot_rule(ws, r, c0, c1):
    for c in range(c0, c1 + 1):
        ws.cell(r, c).border = Border(top=Side('thin', color=TEAL))
        ws.cell(r, c).font = Font(name='Calibri', sz=10, bold=True, color=TEAL)

def body(ws, r, c0, c1, fmt=MONEY, indent=0, bold=False):
    for c in range(c0, c1 + 1):
        x = ws.cell(r, c)
        x.number_format = fmt
        x.font = Font(name='Calibri', sz=10, bold=bold, color=INK)
    ws.cell(r, 1).alignment = Alignment(horizontal='left', indent=indent)

# =========================================================== P&L by Dept ===
ws = sheet('P&L by Dept', 'Overhead Split')
title(ws['A1'], 'P&L by Department')
note(ws['A2'], 'Every line of the P&L with each department down a column. Built straight off '
               'the Cleanup tab, so it moves with the month dropdown and with any new GL paste. '
               'The admin overhead charge comes from the Overhead Split tab.')
ws['A3'] = '=TEXT(Setup!$B$5,"mmm yyyy")'
ws['A3'].font = Font(name='Calibri', sz=11, bold=True, color=TEAL)

NC = len(DEPTS)
COLS = [gl(2 + i) for i in range(NC)] + [gl(2 + NC)]

def grp_block(r, grp, subs, cats, m_lo, m_hi):
    """A sub category block with its own total. Returns the next free row and
    the row the total landed on."""
    band(ws, r, 1, 2 + NC, AQUA)
    ws.cell(r, 1, grp)
    r += 1
    first = r
    for sub in subs:
        ws.cell(r, 1, sub)
        for i, (_, tag) in enumerate(DEPTS):
            parts = '+'.join(
                f'SUMIFS({CA},{CD},"{tag}",{CF},$A{r},{CE},"{c}",'
                f'{CB},">="&{m_lo},{CB},"<="&{m_hi})' for c in cats)
            ws.cell(r, 2 + i, '=' + parts)
        ws.cell(r, 2 + NC, f'=SUM(B{r}:{gl(1+NC)}{r})')
        body(ws, r, 1, 2 + NC, indent=1)
        r += 1
    ws.cell(r, 1, f'Total {grp.lower()}')
    for i in range(NC + 1):
        ws.cell(r, 2 + i, f'=SUM({gl(2+i)}{first}:{gl(2+i)}{r-1})')
    body(ws, r, 1, 2 + NC, bold=True)
    tot_rule(ws, r, 1, 2 + NC)
    return r + 2, r

def line(r, label, formula_for, fmt=MONEY, bold=False, rule=False):
    ws.cell(r, 1, label)
    for i in range(NC + 1):
        ws.cell(r, 2 + i, formula_for(gl(2 + i), i))
    body(ws, r, 1, 2 + NC, fmt=fmt, bold=bold)
    if rule:
        tot_rule(ws, r, 1, 2 + NC)
    return r + 1

def pl_block(r0, label, m_lo, m_hi):
    r = r0
    band(ws, r, 1, 2 + NC, TEAL, WHITE)
    ws.cell(r, 1, label)
    ws.row_dimensions[r].height = 20
    r += 1
    head(ws, r, 1, ['A$'] + [d for d, _ in DEPTS] + ['Total'])
    r += 1
    r, INC = grp_block(r, 'Income', NS['subs']['income'], ['Income'], m_lo, m_hi)
    r, COS = grp_block(r, 'Cost of sales', NS['subs']['cos'], ['Cost of Sales'], m_lo, m_hi)
    GP = r
    r = line(r, 'Gross profit', lambda c, i: f'={c}{INC}+{c}{COS}', bold=True, rule=True)
    r = line(r, 'Gross margin %',
             lambda c, i: f'=IF({c}{INC}=0,"",{c}{GP}/{c}{INC})', fmt=PCT)
    r += 1
    OI = r
    r = line(r, 'Other income', lambda c, i: (
        f'=SUMIFS({CA},{CD},"{DEPTS[i][1]}",{CE},"Other Income",'
        f'{CB},">="&{m_lo},{CB},"<="&{m_hi})' if i < NC
        else f'=SUM(B{OI}:{gl(1+NC)}{OI})'))
    r += 1
    r, OH = grp_block(r, 'Overheads', NS['subs']['exp'] + NS['subs']['oexp'],
                      ['Expenses', 'Other Expenses'], m_lo, m_hi)
    OM = "'Overhead Split'!$C$22:$N$22"
    AL = r
    r = line(r, 'Admin overhead charged across', lambda c, i: (
        f'=SUMIFS(\'Overhead Split\'!$C${90+i}:$N${90+i},{OM},">="&{m_lo},{OM},"<="&{m_hi})'
        if i < 5 else (f'=-SUM(B{AL}:{gl(NC)}{AL})' if i == 5
                       else f'=SUM(B{AL}:{gl(1+NC)}{AL})')))
    RES = r
    r = line(r, 'Result after all overhead',
             lambda c, i: f'={c}{GP}+{c}{OI}+{c}{OH}+{c}{AL}', bold=True, rule=True)
    r = line(r, 'Result %',
             lambda c, i: f'=IF({c}{INC}=0,"",{c}{RES}/{c}{INC})', fmt=PCT)
    return r - 1, {'income': INC, 'cos': COS, 'gp': GP, 'other': OI,
                   'overheads': OH, 'alloc': AL, 'result': RES}

end1, M1 = pl_block(5, 'THE MONTH', MON, MON)
end2, M2 = pl_block(end1 + 3, 'YEAR TO DATE', FY0, MON)

ws.column_dimensions['A'].width = 34
for i in range(NC + 1):
    ws.column_dimensions[gl(2 + i)].width = 14
ws.freeze_panes = 'B6'
json.dump({'month_rows': M1, 'ytd_rows': M2, 'dept_cols': NC}, open(f'{HERE}/data/pldept.json', 'w'))
print(f'P&L by Dept: month block rows 5..{end1}, YTD block {end1+3}..{end2}')

wb.save(OUT)
print('saved', OUT)
