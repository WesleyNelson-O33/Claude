"""
The chargeable target, beside the chargeable per cent.

The deck quoted an 81 per cent target that was typed into the deck code rather
than taken from the monthly report. It is a real figure and it belongs in the
pack, with the over or under reading next to it.

The target is a paste, because it is set per person and averaged up, which
cannot be worked out in here. Production 67 per cent, Consulting 50, Support 88.
The company figure moves month to month because it is an average over whoever
was in the report that month.

Over or under target, and leave as a share of available hours, are calculated
rather than pasted. Both are plain subtraction and division off figures already
on the tab, and both reproduce the report's own columns exactly.

OT/TIL is deliberately left out. The report's August figure for Support reads
365.62, which cannot be hours in a month, so that column looks wrong on the
source sheet. I am not putting a number in the pack that I cannot stand behind.

This rebuilds rows 120 to 163, replacing the shorter block from patch_v54.
"""
import json, os, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
WB = sys.argv[1] if len(sys.argv) > 1 else f'{HERE}/CTS Financial Controller Pack - v20.xlsx'
H = json.load(open(f'{HERE}/../deck/data/util_her_basis.json'))

NAVY, WHITE, GREYL = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTAL = PatternFill('solid', start_color=GREYL, end_color=GREYL)
NOFILL = PatternFill(fill_type=None)
RULE = Side(style='medium', color=NAVY)
HRS, PCT = '#,##0.00;;', '0.0%'
SIGNED = '0.0%;[RED]-0.0%;\\-'

#            label                                 key        kind
LINES = [('Billable hours',                        'ch',      'paste'),
         ('Non-billable hours',                    'nc',      'paste'),
         ('Leave',                                 'lv',      'paste'),
         ('Public holiday',                        'ph',      'paste'),
         ('Workable business hours',               'workable', 'paste'),
         ('Chargeable %',                          None,      'pct'),
         ('Chargeable target',                     'target',  'paste'),
         ('Over or under target',                  None,      'over'),
         ('Leave, share of available hours',       None,      'leavepct')]
GROUPS = [('Production', 123), ('Consulting', 134), ('Support', 145)]
COMPANY = 156
MONTHCOL = {'Jul-26': 4, 'Aug-26': 5, 'Sep-26': 6}      # D, E, F

wb = openpyxl.load_workbook(WB)
ws = wb['Utilisation']

# clear what patch_v54 put down, so the longer block can be laid out clean
for r in range(120, 170):
    for c in range(1, 16):
        cell = ws.cell(r, c)
        cell.value = None
        cell.fill = NOFILL
        cell.border = Border()

ws['A120'] = 'UTILISATION, AS THE MONTHLY REPORT CALCULATES IT'
for c in range(1, 16):
    cell = ws.cell(120, c)
    cell.fill = BAND
    cell.font = Font(name='Calibri', sz=12, b=True, color=WHITE)

for c, text in enumerate(['FY (year it ends)', 'Person Group', 'Hours'], start=1):
    cell = ws.cell(122, c, text)
    cell.fill = BAND
    cell.font = Font(name='Calibri', sz=10, b=True, color=WHITE)
    cell.border = Border(bottom=RULE)
for c in range(4, 16):
    cell = ws.cell(122, c, f'={gl(c)}$4')
    cell.number_format = 'mmm-yy'
    cell.fill = BAND
    cell.font = Font(name='Calibri', sz=10, b=True, color=WHITE)
    cell.border = Border(bottom=RULE)
    cell.alignment = Alignment(horizontal='right')

def block(r0, name, vals):
    """vals is None for the company total, which sums the three groups."""
    for i, (label, key, kind) in enumerate(LINES):
        r = r0 + i
        if i == 0:
            ws.cell(r, 1, '=Setup!$B$7')
            ws.cell(r, 2, name).font = Font(name='Calibri', sz=11, b=True, color=NAVY)
        bold = kind in ('pct', 'over')
        ws.cell(r, 3, label).font = Font(name='Calibri', sz=11, b=bold, color=NAVY)
        for c in range(4, 16):
            col, cell = gl(c), ws.cell(r, c)
            if kind == 'pct':
                cell.value = f'=IFERROR({col}{r0}/{col}{r0 + 4},0)'
                cell.number_format = PCT
            elif kind == 'over':
                cell.value = (f'=IF({col}{r0 + 6}="","",{col}{r0 + 5}-{col}{r0 + 6})')
                cell.number_format = SIGNED
            elif kind == 'leavepct':
                cell.value = f'=IFERROR({col}{r0 + 2}/{col}{r0 + 4},0)'
                cell.number_format = PCT
            elif vals is None:
                # the company line sums the three groups, except the target,
                # which is an average over people and so is pasted
                if key == 'target':
                    mon = next((m for m, cc in MONTHCOL.items() if cc == c), None)
                    cell.value = H[mon]['Grand']['target'] if mon else None
                    cell.number_format = PCT
                    cell.font = Font(name='Calibri', sz=11, color=NAVY)
                    continue
                cell.value = '=' + '+'.join(f'{col}{g[1] + i}' for g in GROUPS)
                cell.number_format = HRS
            else:
                mon = next((m for m, cc in MONTHCOL.items() if cc == c), None)
                cell.value = vals[key][mon] if mon else None
                cell.number_format = PCT if key == 'target' else HRS
            cell.font = Font(name='Calibri', sz=11, b=bold, color=NAVY)

KEYS = [k for _, k, _ in LINES if k]
for name, r0 in GROUPS:
    block(r0, name, {k: {m: H[m][name][k] for m in MONTHCOL} for k in KEYS})
block(COMPANY, 'Company', None)
for i in range(len(LINES)):
    for c in range(1, 16):
        ws.cell(COMPANY + i, c).fill = TOTAL

# Against the Grand Total rows of the report itself, each month.
for r, label, key in ((166, 'Check, the report\'s Grand Total chargeable %', 'pct'),
                      (167, 'Check, the report\'s Grand Total over or under', 'over_under')):
    ws.cell(r, 3, label).font = Font(name='Calibri', sz=11, b=True, color=NAVY)
    for m, c in MONTHCOL.items():
        cell = ws.cell(r, c, round(H[m]['Grand'][key], 6))
        cell.number_format = SIGNED if key == 'over_under' else PCT
        cell.font = Font(name='Calibri', sz=11, b=True, color=NAVY)

wb.save(WB)
print('Utilisation rows 120 to 167 rebuilt: target, over or under, leave share')
