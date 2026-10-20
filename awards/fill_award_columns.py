"""Fill the award columns on the CTS Payroll Setup Register (21/10/2026).

Keeps everything already in Danica's file. Writes only the yellow columns
V (Award) and W (Classification), plus a new tick column AA.
"""
import sys
from copy import copy
import openpyxl

src, out = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src)
ws = wb['Register']

LPA = 'Live Performance Award 2020 [MA000081]'
CLERKS = 'Clerks Private Sector Award 2020 [MA000002]'
TICK = '✓'

# Not award covered on current setup: no agreement in EH, senior management roles.
award_free = {
    ('Danica', 'Nelson'), ('Duncan', 'LUGSTEIN'), ('Gianmaria', 'Yao'),
    ('Graham', 'Cherry'), ('Jordan', 'SEXTY'),
}
clerks = {('Emelda', 'Waiya'), ('Isabella', 'NIKODINOVSKI')}
check_accountant = {('Premila', 'LUGSTEIN')}
# On an LPA level in EH, but the role may sit outside live performance work.
lpa_check = {
    ('Daniel', 'SOBKOWSKI'): 'Integration role. Check the duties are AV production work and not installation, which may fall under the Electrical award [MA000025].',
    ('Mitchell', 'Healey'): 'Production Manager. May be award free if the role is truly managerial. Contract decides.',
    ('Tyler', 'WOOD'): 'Senior project manager. May be award free if the role is truly managerial. Contract decides.',
    ('Simon', 'GRUENEFELD'): 'Account manager title. Check the duties are production and support work.',
    ('Patrick', 'LEE'): 'Workplace technology role. Check the duties are production and support work.',
}

# New tick column AA, styled like column Z.
hdr = ws['AA1']
hdr.value = 'Under an award? (tick)'
for r in range(1, 81):
    s, d = ws[f'Z{r}'], ws[f'AA{r}']
    d.font, d.fill, d.border, d.alignment = copy(s.font), copy(s.fill), copy(s.border), copy(s.alignment)
    d.number_format = s.number_format
ws.column_dimensions['AA'].width = 14
ws.column_dimensions['V'].width = 34
ws.column_dimensions['W'].width = 46
ws.auto_filter.ref = 'A1:AA80'

for r in range(2, 81):
    key = (ws[f'A{r}'].value, ws[f'B{r}'].value)
    agreement = ws[f'D{r}'].value or ''
    if key in award_free:
        ws[f'V{r}'] = 'Award free (no agreement in EH)'
        ws[f'W{r}'] = 'Senior or salaried role. National Employment Standards and minimum wage still apply.'
        ws[f'AA{r}'] = ''
    elif key in clerks:
        ws[f'V{r}'] = CLERKS
        ws[f'W{r}'] = 'Clerks level 1 to 5 to be set from duties. Currently on LPA ' + agreement.replace(' - ', ' ')
        ws[f'AA{r}'] = TICK
    elif key in check_accountant:
        ws[f'V{r}'] = 'CHECK: ' + CLERKS + ' or award free'
        ws[f'W{r}'] = 'Clerks award if the work is accounts clerk work. Award free if it is professional accounting work. Contract and duties decide.'
        ws[f'AA{r}'] = TICK + ' check'
    elif agreement.startswith('Production & Support Staff - Level'):
        lvl = agreement.split('Level')[-1].strip()
        ws[f'V{r}'] = LPA if key not in lpa_check else 'CHECK: ' + LPA
        ws[f'W{r}'] = f'Production and Support Staff Level {lvl} (as set up in EH)'
        if key in lpa_check:
            ws[f'W{r}'] = ws[f'W{r}'].value + '. ' + lpa_check[key]
        ws[f'AA{r}'] = TICK if key not in lpa_check else TICK + ' check'
    else:
        raise SystemExit(f'Unhandled row {r}: {key} {agreement}')

# Award References: add the award the EH levels actually come from. Nothing removed.
ar = wb['Award References']
row = ar.max_row + 1
for c in (1, 2, 3):
    s = ar.cell(6, c)
    d = ar.cell(row, c)
    d.font, d.alignment, d.border = copy(s.font), copy(s.alignment), copy(s.border)
ar.cell(row, 1, 'Live Performance Award 2020 [MA000081]')
ar.cell(row, 2, 'The award CTS already uses. The EH agreements "Production & Support Staff - Level 1 to 8" are this award\'s classification names (clause 11.1, Schedule A). Casual production staff: Part 8, clauses 55 to 63. Overtime and penalties: clause 63. Rates from first pay period on or after 1 July 2026: Level 3 $28.89, Level 4 $29.45, Level 5 $30.37, Level 6 $31.30, Level 7 $33.311 per hour, before the 25% casual loading.')
ar.cell(row, 3, 'https://awards.fairwork.gov.au/MA000081.html')
wb.save(out)
print('saved', out)
