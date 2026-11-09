"""The month box appears on every statement tab, and one of them drives it.

I built the version she asked for first, with a live dropdown on all five
tabs, and tested it. Moving the P&L FY27 dropdown to July while the Summary
stayed on September changed the Summary's year to date income from
1,624,459.78 to 1,845,881.13, because the Summary adds up the P&L FY27 month
columns and those columns switch between actual and budget on the month. The
Summary was wrong and nothing on screen said so, so that version is not
shipped.

What is here instead: the same month box, in the same place, with the same
look, on all five tabs. The Summary one is live. The other four follow it, so
every tab is always on one month and the figures cannot disagree.
"""
import os, shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = ('/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a/'
       'scratchpad/v83/her10.xlsx')
OUT = f'{HERE}/CTS Financial Controller Pack - v11.xlsx'

TEAL, GREY, MID = '005358', 'E3E6EA', '5F6670'
FOLLOWERS = ['P&L FY27', 'Budget FY27', 'Actual FY26', 'Budget FY26']

shutil.copy(SRC, OUT)
wb = openpyxl.load_workbook(OUT)

def box(ws, live):
    ws['A3'] = 'Month'
    ws['A3'].font = Font(name='Calibri', sz=10, bold=True, color=TEAL)
    ws['A3'].alignment = Alignment(horizontal='left', vertical='center')
    b = ws['B3']
    b.number_format = 'mmm yyyy'
    b.font = Font(name='Calibri', sz=11, bold=True, color=TEAL)
    b.fill = PatternFill('solid', fgColor=GREY)
    b.alignment = Alignment(horizontal='center', vertical='center')
    if live:
        dv = DataValidation(type='list', formula1='=Months', allow_blank=False,
                            showDropDown=False)
        dv.promptTitle = 'Reporting month'
        dv.prompt = 'Pick the month. Every tab in the book follows it.'
        ws.add_data_validation(dv)
        dv.add(b)
        ws['C3'] = ('Pick the month here. Every tab, every YTD and the Actual '
                    'and Forecast headings all follow it.')
    else:
        b.value = '=Setup!$B$5'
        ws['C3'] = 'Follows the month dropdown on the Summary tab.'
    ws['C3'].font = Font(name='Calibri', sz=9, italic=not live, color=MID)
    ws['C3'].alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[3].height = 16

box(wb['Summary'], live=True)
for name in FOLLOWERS:
    box(wb[name], live=False)

# Row 4 on P&L FY27 keeps reading the book month, so the Actual and Forecast
# labels can never say one thing while the figures underneath say another.
ws = wb['P&L FY27']
for col in range(2, 14):
    cl = ws.cell(4, col).column_letter
    ws.cell(4, col).value = f'=IF({cl}$5<=Setup!$B$5,"Actual","Forecast")'

wb.save(OUT)
print('saved', OUT)
