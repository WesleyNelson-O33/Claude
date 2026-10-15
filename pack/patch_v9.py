"""Work on her own v8_3 file. Nothing is rebuilt, nothing of hers is replaced.

Three changes, in her words:
  1. the minus amounts go red on their own
  2. a month dropdown that drives the book, and a row above the month headings
     that reads Actual for a month that is done and Forecast for one that is not
  3. everything, YTD included, follows the month she picks
"""
import os, shutil, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = ('/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a/'
       'scratchpad/v83/her.xlsx')
OUT = f'{HERE}/CTS Financial Controller Pack - v9.xlsx'

TEAL, AQUA, GREY, GREY_LT = '005358', '3FD0C9', 'E3E6EA', 'F4F4F4'
MID, INK, WHITE = '5F6670', '1A1A1A', 'FFFFFF'

shutil.copy(SRC, OUT)
wb = openpyxl.load_workbook(OUT)

# --------------------------------------------------------------- 1. red ----
# A number format has up to three parts: positive;negative;zero. [Red] at the
# front of the middle part turns every negative figure red by itself. No
# conditional formatting, so there is nothing to go stale and it survives a
# paste and a recalculation.
def red(fmt):
    if not fmt or fmt == 'General' or '[Red]' in fmt or '[RED]' in fmt:
        return fmt
    parts = fmt.split(';')
    if len(parts) < 2:
        return fmt
    parts[1] = '[Red]' + parts[1]
    return ';'.join(parts)

n = 0
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            new = red(c.number_format)
            if new != c.number_format:
                c.number_format = new
                n += 1
print('negatives set to show red:', n)

# ------------------------------------------------ 2. the month dropdown ----
# Excel has no way to let the same value be typed into twelve places at once,
# so there is one live dropdown and every other tab shows what it is set to.
# It sits on Summary, where she works. Setup row 5 follows it, so every
# formula in the book that already read Setup!$B$5 keeps working untouched.
SUM_WS = wb['Summary']
SUM_WS['A3'] = 'Month'
SUM_WS['A3'].font = Font(name='Calibri', sz=10, bold=True, color=TEAL)
SUM_WS['A3'].alignment = Alignment(horizontal='left', vertical='center')

sel = SUM_WS['B3']
if not isinstance(sel.value, (int, float)) and sel.value is None:
    sel.value = wb['Setup']['B5'].value        # start on the month she is on
sel.number_format = 'mmm yyyy'
sel.font = Font(name='Calibri', sz=11, bold=True, color=TEAL)
sel.fill = PatternFill('solid', fgColor=GREY)
sel.alignment = Alignment(horizontal='center', vertical='center')

dv = DataValidation(type='list', formula1='=Months', allow_blank=False,
                    showDropDown=False)
dv.prompt = 'Pick the month you are reporting on. The whole book follows it.'
dv.promptTitle = 'Reporting month'
SUM_WS.add_data_validation(dv)
dv.add(sel)

SUM_WS['C3'] = 'Pick the month here. Every tab, every YTD and the Actual and Forecast headings all follow it.'
SUM_WS['C3'].font = Font(name='Calibri', sz=9, italic=True, color=MID)
SUM_WS['A4'] = None                            # her note, now done

# Setup keeps its row, it just follows the dropdown instead of being typed in.
ST = wb['Setup']
ST['B5'] = '=Summary!$B$3'
ST['B5'].number_format = 'mmm yyyy'
ST['C5'] = 'Follows the month dropdown on the Summary tab.'
ST['A2'] = ('The month is set by the dropdown on the Summary tab. '
            'The other cells in blue text are the only ones you type into.')
for dvv in list(ST.data_validations.dataValidation):
    if 'B5' in str(dvv.sqref):
        ST.data_validations.dataValidation.remove(dvv)

# ------------------------------- 3. Actual or Forecast above each month ----
# P&L FY27 is actual up to the chosen month and budget after it, so its
# heading really does change. The budget and prior year tabs do not switch,
# so they say what they are rather than pretending to.
SWITCHING = {'P&L FY27': (2, 13), 'Summary': (2, 4)}
FIXED = {'Budget FY27': ('Budget', 2, 13),
         'Actual FY26': ('Actual', 2, 13),
         'Budget FY26': ('Budget', 2, 13)}

def band(ws, c0, c1):
    for c in range(c0, c1 + 1):
        cell = ws.cell(4, c)
        cell.font = Font(name='Calibri', sz=9, bold=True, color=TEAL)
        cell.fill = PatternFill('solid', fgColor=GREY_LT)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.number_format = 'General'
    ws.row_dimensions[4].height = 14

for nm, (c0, c1) in SWITCHING.items():
    ws = wb[nm]
    for c in range(c0, c1 + 1):
        col = ws.cell(4, c).column_letter
        ws.cell(4, c).value = f'=IF({col}$5<=Setup!$B$5,"Actual","Forecast")'
    band(ws, c0, c1)

for nm, (word, c0, c1) in FIXED.items():
    ws = wb[nm]
    for c in range(c0, c1 + 1):
        ws.cell(4, c).value = word
    band(ws, c0, c1)

# The month a tab is showing, printed on every tab that has no dropdown.
for nm in ['P&L FY27', 'Budget FY27', 'Actual FY26', 'Budget FY26',
           'Balance_Sheet', 'PL_Check', 'Utilisation']:
    ws = wb[nm]
    ws['A3'] = '=TEXT(Setup!$B$5,"mmm yyyy")&"   |   month is set on the Summary tab"'
    ws['A3'].font = Font(name='Calibri', sz=9, italic=True, color=MID)
    ws['A3'].alignment = Alignment(horizontal='left', vertical='center')

wb.save(OUT)
print('saved', OUT)
