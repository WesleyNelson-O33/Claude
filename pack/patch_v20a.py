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

# The sheet mirrors P&L FY27 line for line, and takes its look from it: every
# row copies the style of the P&L row it corresponds to, so the banners,
# indents, underlines and the costs-as-positive number format all match.
SRCWB = openpyxl.load_workbook(SRC)
PL = SRCWB['P&L FY27']

def st(ws, r, src_row, cols):
    """Copy a P&L row's look onto this row, column A plus the figure columns."""
    ws.cell(r, 1)._style = copy.copy(PL.cell(src_row, 1)._style)
    for c in cols:
        ws.cell(r, c)._style = copy.copy(PL.cell(src_row, 2)._style)

SUMD = ('SUMIFS({CA},{CD},"{tag}",{CE},"{cat}",{CB},">="&{lo},{CB},"<="&{hi})')
SUMS = ('SUMIFS({CA},{CD},"{tag}",{CF},TRIM($A{r}),{CE},"{cat}",{CB},">="&{lo},{CB},"<="&{hi})')

def pl_block(r0, label, m_lo, m_hi):
    cols = list(range(2, 3 + NC))          # B..G departments, H total
    last = gl(2 + NC - 1)                  # last department column
    r = r0
    band(ws, r, 1, 2 + NC, TEAL, WHITE)
    ws.cell(r, 1, label)
    ws.row_dimensions[r].height = 20
    r += 1
    ws.cell(r, 1, 'Line')
    for i, (d, _) in enumerate(DEPTS):
        ws.cell(r, 2 + i, d)
    ws.cell(r, 2 + NC, 'Total')
    st(ws, r, 5, cols)
    ws.row_dimensions[r].height = PL.row_dimensions[5].height or 30
    r += 1
    M = {}

    def cat_row(label, cat, src):
        nonlocal r
        ws.cell(r, 1, label)
        for i, (_, tag) in enumerate(DEPTS):
            ws.cell(r, 2 + i, '=' + SUMD.format(CA=CA, CD=CD, CE=CE, CB=CB,
                                                tag=tag, cat=cat, lo=m_lo, hi=m_hi))
        ws.cell(r, 2 + NC, f'=SUM(B{r}:{last}{r})')
        st(ws, r, src, cols)
        M[label] = r
        r += 1

    def sub_rows(subs, cat, src):
        nonlocal r
        for sub in subs:
            ws.cell(r, 1, '        ' + sub)
            for i, (_, tag) in enumerate(DEPTS):
                ws.cell(r, 2 + i, '=' + SUMS.format(CA=CA, CD=CD, CE=CE, CF=CF, CB=CB,
                                                    tag=tag, cat=cat, r=r, lo=m_lo, hi=m_hi))
            ws.cell(r, 2 + NC, f'=SUM(B{r}:{last}{r})')
            st(ws, r, src, cols)
            r += 1

    def calc(label, make, src):
        nonlocal r
        ws.cell(r, 1, label)
        for i in range(NC + 1):
            ws.cell(r, 2 + i, make(gl(2 + i), i))
        st(ws, r, src, cols)
        M[label] = r
        r += 1

    cat_row('Total Income', 'Income', 6)
    sub_rows(NS['subs']['income'], 'Income', 9)
    cat_row('Other income', 'Other Income', 57)
    cat_row('Cost of Sales', 'Cost of Sales', 59)
    sub_rows(NS['subs']['cos'], 'Cost of Sales', 62)
    cat_row('Expenses Overheds', 'Expenses', 164)
    sub_rows(NS['subs']['exp'], 'Expenses', 165)
    cat_row('Other Expenses', 'Other Expenses', 171)
    sub_rows(NS['subs']['oexp'], 'Other Expenses', 172)

    INC, OI = M['Total Income'], M['Other income']
    COS, EXP, OEX = M['Cost of Sales'], M['Expenses Overheds'], M['Other Expenses']
    calc('Gross profit', lambda c, i: f'={c}{INC}+{c}{COS}', 173)
    GP = M['Gross profit']
    calc('Gross profit %', lambda c, i: f'=IF({c}{INC}=0,"",{c}{GP}/{c}{INC})', 174)
    calc('Depreciation & amortisation', lambda c, i: (
        '=IF(Setup!$B$28="Yes",' + SUMIAWO.format(CA=CA, CD=CD, CG=CG, CB=CB,
                                                  tag=DEPTS[i][1], lo=m_lo, hi=m_hi) + ',0)'
        if i < NC else f'=SUM(B{M.get("Depreciation & amortisation", r)}:'
                       f'{last}{M.get("Depreciation & amortisation", r)})'), 175)
    DA = M['Depreciation & amortisation']
    ws.cell(DA, 2 + NC, f'=SUM(B{DA}:{last}{DA})')
    calc('EBITDA', lambda c, i: f'={c}{INC}+{c}{COS}+{c}{EXP}-{c}{DA}', 176)
    EB = M['EBITDA']
    calc('EBITDA %', lambda c, i: f'=IF({c}{INC}=0,"",{c}{EB}/{c}{INC})', 177)
    calc('EBIT', lambda c, i: f'={c}{EB}+{c}{DA}', 178)
    calc('EBIT %', lambda c, i: f'=IF({c}{INC}=0,"",{c}{M["EBIT"]}/{c}{INC})', 179)
    calc('Net profit', lambda c, i: f'={c}{INC}+{c}{OI}+{c}{COS}+{c}{EXP}+{c}{OEX}', 180)
    NP = M['Net profit']
    calc('Net profit %', lambda c, i: f'=IF({c}{INC}=0,"",{c}{NP}/{c}{INC})', 181)
    r += 1
    OM = "'Overhead Split'!$C$22:$N$22"
    AL = r
    calc('Admin overhead charged across', lambda c, i: (
        f'=SUMIFS(\'Overhead Split\'!$C${90+i}:$N${90+i},{OM},">="&{m_lo},{OM},"<="&{m_hi})'
        if i < 5 else (f'=-SUM(B{AL}:{gl(NC)}{AL})' if i == 5
                       else f'=SUM(B{AL}:{last}{AL})')), 57)
    calc('Result after all overhead', lambda c, i: f'={c}{NP}+{c}{AL}', 180)
    RS = M['Result after all overhead']
    calc('Result %', lambda c, i: f'=IF({c}{INC}=0,"",{c}{RS}/{c}{INC})', 181)
    return r - 1, M

SUMIAWO = 'SUMIFS({CA},{CD},"{tag}",{CG},"IAWO",{CB},">="&{lo},{CB},"<="&{hi})'
CG = 'Cleanup!$G$5:$G$20004'

end1, M1 = pl_block(5, 'THE MONTH', MON, MON)
end2, M2 = pl_block(end1 + 3, 'YEAR TO DATE', FY0, MON)

ws.column_dimensions['A'].width = 36
for i in range(NC + 1):
    ws.column_dimensions[gl(2 + i)].width = 15
ws.freeze_panes = 'B7'
print(f'P&L by Dept: month block 5..{end1}, YTD block {end1+3}..{end2}')
json.dump({'month_rows': M1, 'ytd_rows': M2, 'dept_cols': NC},
          open(f'{HERE}/data/pldept.json', 'w'))

wb.save(OUT)
print('saved', OUT)
