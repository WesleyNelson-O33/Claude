"""Add a Casual Rate Check tab to the payroll setup register.

Compares each casual's EH rate (Register column G) with the Casual Ordinary
Hours rate EH paid in the pay run audit, and with the Live Performance Award
base rate for their level (from first pay period on or after 1 July 2026).
"""
import sys
from copy import copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

reg_path, audit_path = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(reg_path)
reg = wb['Register']
audit = openpyxl.load_workbook(audit_path, data_only=True)['Earnings Details']

paid = {}
for r in audit.iter_rows(min_row=3, values_only=True):
    if r[2] and r[5] == 'Casual Ordinary Hours':
        paid.setdefault(r[2].strip().lower(), set()).add(round(r[12], 4))

if 'Casual Rate Check' in wb.sheetnames:
    del wb['Casual Rate Check']
ws = wb.create_sheet('Casual Rate Check', index=wb.sheetnames.index('Register') + 1)
F = 'Arial'
hdr_fill = copy(reg['A1'].fill)
ws['A1'] = 'Casual rate check: pay run 8, fortnight ending 16/10/2026'
ws['A1'].font = Font(name=F, bold=True, size=12)
ws['A2'] = ('Source: EH Pay Run Audit report, Earnings Details tab, "Casual Ordinary Hours" lines. '
            'Award base rates: Live Performance Award 2020 [MA000081] clause 11.1, from first pay period on or after 1 July 2026.')
ws['A2'].font = Font(name=F, size=9, italic=True)
ws['A3'] = 'Award base rate per hour'
ws['A3'].font = Font(name=F, bold=True)
mins = {3: 28.89, 4: 29.45, 5: 30.37, 6: 31.30, 7: 33.311}
for i, (lvl, v) in enumerate(mins.items()):
    ws.cell(4 + i, 1, f'Level {lvl}').font = Font(name=F)
    c = ws.cell(4 + i, 2, v); c.font = Font(name=F, color='0000FF'); c.number_format = '$#,##0.00##'
hr = 10
heads = ['First name', 'Surname', 'Level', 'EH rate (base, Register G)', 'Rate EH paid in pay run 8',
         'Loading check (paid = base x 1.25)', 'Award base rate', 'Award check', 'Result']
for j, h in enumerate(heads, 1):
    c = ws.cell(hr, j, h)
    c.font = Font(name=F, bold=True, color='FFFFFF'); c.fill = copy(hdr_fill)
    c.alignment = Alignment(wrap_text=True, vertical='top')
row = hr + 1
red = PatternFill('solid', fgColor='FFC7CE')
for i in range(2, 81):
    if reg[f'C{i}'].value != 'Casual':
        continue
    fn, sn = reg[f'A{i}'].value, reg[f'B{i}'].value
    lvl = int(str(reg[f'D{i}'].value).strip()[-1])
    p = paid.get(f'{fn} {sn}'.lower())
    ws.cell(row, 1, fn); ws.cell(row, 2, sn); ws.cell(row, 3, lvl)
    ws.cell(row, 4, f'=Register!G{i}')
    if p:
        if len(p) > 1:
            raise SystemExit(f'{fn} {sn} has more than one ordinary rate: {p}')
        ws.cell(row, 5, p.pop()).font = Font(name=F, color='0000FF')
        ws.cell(row, 6, f'=IF(ABS(E{row}-D{row}*1.25)<0.01,"Loading added","CHECK")')
    else:
        ws.cell(row, 5, 'Not in pay run 8')
        ws.cell(row, 6, 'Not tested')
    ws.cell(row, 7, f'=INDEX($B$4:$B$8,MATCH("Level "&C{row},$A$4:$A$8,0))')
    ws.cell(row, 8, f'=IF(D{row}<G{row}-0.005,"BELOW AWARD","OK")')
    ws.cell(row, 9, f'=IF(H{row}="BELOW AWARD","Underpaid. Raise EH rate to at least "&TEXT(G{row},"$0.00")&" and back pay",IF(F{row}="CHECK","Loading not applied","OK"))')
    for j in range(1, 10):
        c = ws.cell(row, j)
        c.font = Font(name=F, color=c.font.color.rgb if j == 5 and isinstance(c.value, float) else None)
        if j == 5 and isinstance(c.value, float):
            c.font = Font(name=F, color='0000FF')
        if j in (4, 5, 7) and not (j == 5 and isinstance(c.value, str)):
            c.number_format = '$#,##0.00##'
    row += 1
last = row - 1
from openpyxl.formatting.rule import FormulaRule
ws.conditional_formatting.add(f'A{hr+1}:I{last}', FormulaRule(formula=[f'$H{hr+1}="BELOW AWARD"'], fill=red))
ws.conditional_formatting.add(f'A{hr+1}:I{last}', FormulaRule(formula=[f'$F{hr+1}="CHECK"'], fill=red))
ws.cell(last + 2, 1, 'Casuals paid in pay run 8').font = Font(name=F, bold=True)
ws.cell(last + 2, 4, f'=COUNTIF(F{hr+1}:F{last},"Loading added")+COUNTIF(F{hr+1}:F{last},"CHECK")').font = Font(name=F)
ws.cell(last + 3, 1, 'Loading added correctly').font = Font(name=F, bold=True)
ws.cell(last + 3, 4, f'=COUNTIF(F{hr+1}:F{last},"Loading added")').font = Font(name=F)
ws.cell(last + 4, 1, 'Below award base rate').font = Font(name=F, bold=True)
ws.cell(last + 4, 4, f'=COUNTIF(H{hr+1}:H{last},"BELOW AWARD")').font = Font(name=F)
for col, w in zip('ABCDEFGHI', (14, 22, 7, 14, 14, 16, 12, 13, 52)):
    ws.column_dimensions[col].width = w
ws.freeze_panes = f'A{hr+1}'
ws.auto_filter.ref = f'A{hr}:I{last}'
wb.save(reg_path)
print('rows', hr + 1, last)
