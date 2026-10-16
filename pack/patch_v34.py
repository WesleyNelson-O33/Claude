"""The employment expense split on the Summary too.

Office staff is the admin pot, the employment expense sitting on the CTS cost
centre. Departmental staff is everything else, which is the five departments'
own. Added as a block under the statement rather than as rows inside it,
because inserting rows into that block would shift 50 rows that the Engine,
P&L FY27, Actual FY26, Budget FY26 and P&L by Dept all reference by number.
"""
import os, json, copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = f'{HERE}/CTS Financial Controller Pack - v20.xlsx'
B = json.load(open(f'{HERE}/data/brand.json'))
STAFF = json.load(open(f'{HERE}/data/staffsplit.json'))
POT = STAFF['pot']
INK, WHITE, GREY, GREYL, AQUA = (B['charcoal'], B['white'], B['grey'],
                                 B['grey_lt'], B['bright_teal'])
COST = '[Red]"($"#,##0\\);\\$#,##0;\\-'
OMM = "'Overheads MoM'"
EMP = 169                                  # the Employment Expenses row
S = 194
thin = Side(style='thin', color=GREY)
hair = Side(style='medium', color=AQUA)

wb = openpyxl.load_workbook(OUT)
s = wb['Summary']

s.cell(S, 1, 'Employment expenses, office staff and departmental staff')
for c in range(1, 6):
    s.cell(S, c).fill = PatternFill('solid', fgColor=INK)
    s.cell(S, c).font = Font(name='Calibri', size=11, bold=True, color=WHITE)
    s.cell(S, c).border = Border(bottom=hair)

def row(r, label, body, bold=False, fill=None):
    s.cell(r, 1, label)
    for c in (2, 3, 4, 5):
        s.cell(r, c, body(c))
        s.cell(r, c).number_format = COST
    for c in range(1, 6):
        s.cell(r, c).font = Font(name='Calibri', size=11, bold=bold, color=INK)
        s.cell(r, c).border = Border(bottom=thin)
        if fill:
            s.cell(r, c).fill = PatternFill('solid', fgColor=fill)
        if c > 1:
            s.cell(r, c).alignment = Alignment(horizontal='right')

pick = lambda c: (f'=SUMIF({OMM}!$B$4:$M$4,"<="&Setup!$B$5,{OMM}!$B${POT}:$M${POT})'
                  if c == 5 else
                  f'=SUMIF({OMM}!$B$4:$M$4,{openpyxl.utils.get_column_letter(c)}$5,'
                  f'{OMM}!$B${POT}:$M${POT})')
row(S + 1, '    Office staff, the admin pot', pick)
row(S + 2, '    Departmental staff',
    lambda c: f'={openpyxl.utils.get_column_letter(c)}{EMP}'
              f'-{openpyxl.utils.get_column_letter(c)}{S+1}')
row(S + 3, '    Employment expenses in total',
    lambda c: f'={openpyxl.utils.get_column_letter(c)}{S+1}'
              f'+{openpyxl.utils.get_column_letter(c)}{S+2}', bold=True, fill=GREYL)
row(S + 4, '    Check, against the employment expense line above',
    lambda c: f'=ROUND({openpyxl.utils.get_column_letter(c)}{S+3}'
              f'-{openpyxl.utils.get_column_letter(c)}{EMP},2)', bold=True)
for r in range(S, S + 5):
    s.row_dimensions[r].outlineLevel = 0
    s.row_dimensions[r].hidden = False
print(f'Summary: split block rows {S} to {S+4}')
wb.save(OUT)
print('saved')
