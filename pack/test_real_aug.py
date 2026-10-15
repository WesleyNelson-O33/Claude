"""Paste Danica's real August Xero export into the pack and tie it to Xero."""
import json, shutil, subprocess, sys, datetime as dt
import openpyxl
HERE = __file__.rsplit('/', 1)[0]
D = '/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a/scratchpad/aug'
SRC, TST = f'{HERE}/CTS Financial Controller Pack - v5.xlsx', f'{HERE}/_real.xlsx'
rows = json.load(open(f'{D}/gl_aug.json'))

shutil.copy(SRC, TST)
wb = openpyxl.load_workbook(TST)
gp = wb['P&L Analysis']
for i, row in enumerate(rows, start=1):                 # whole export, title rows and all, at A1
    for c, v in enumerate(row[:12], start=1):           # the twelve Xero columns
        if v is None: continue
        if c == 4 and isinstance(v, str) and len(v) >= 10:
            try: v = dt.datetime.fromisoformat(v)
            except ValueError: pass
        gp.cell(i, c, v)
wb['Setup']['B5'] = dt.datetime(2026, 8, 1)             # report on August
wb.save(TST)
print(f'pasted {len(rows)} rows of her export, reporting month August 2026')
j = json.loads(subprocess.run([sys.executable, '/mnt/skills/public/xlsx/scripts/recalc.py',
                               TST, '1800'], capture_output=True, text=True).stdout)
print('recalc:', j['status'], j['total_errors'], 'errors')

v = openpyxl.load_workbook(TST, data_only=True)
st, pl = v['Setup'], v['P&L FY27']
print()
print('--- what the pack read ---')
for r, lab in ((20, 'GL lines pasted'), (32 - 12, ''), (22, 'no department tag'),
               (23, 'account not on Lists'), (24, 'earliest date'), (25, 'latest date'),
               (36, 'dated outside this FY')):
    if not lab: continue
    print(f'   {lab:<24}{st.cell(r, 2).value}')
LBL = {pl.cell(r, 1).value: r for r in range(6, 190) if pl.cell(r, 1).value}
# Her export is the target: the pack must reproduce what she pasted, to the cent.
HER = {'Total Income ': 603466.54, '    Cost of Sales': -335516.00,
       'Other income ': 861.35}
HER_EXP, HER_NET = -208889.82, 59922.07
XERO_EXP, XERO_NET = -208876.19, 59935.70
ok = True
def chk(lab, got, want):
    global ok
    good = abs(round(got or 0, 2) - round(want, 2)) < 0.02
    ok = ok and good
    print(f'   {lab:<36}{(got or 0):>14,.2f} vs {want:>14,.2f}   '
          f'{"OK" if good else "OUT BY %.2f" % ((got or 0) - want)}')
AUG = 3
print('--- August: the pack against the export she pasted ---')
for lab, want in HER.items(): chk(lab.strip(), pl.cell(LBL[lab], AUG).value, want)
exp = (pl.cell(LBL['    Expenses Overheds '], AUG).value or 0) \
      + (pl.cell(LBL['    Other Expenses'], AUG).value or 0)
chk('Expenses incl interest and tax', exp, HER_EXP)
chk('Net profit', pl.cell(6, AUG).value, HER_NET)
print()
print('--- the same export against live Xero ---')
print(f'   Expenses   her export {HER_EXP:>14,.2f}   live Xero {XERO_EXP:>14,.2f}   '
      f'differs by {HER_EXP - XERO_EXP:>8,.2f}')
print(f'   Net profit her export {HER_NET:>14,.2f}   live Xero {XERO_NET:>14,.2f}   '
      f'differs by {HER_NET - XERO_NET:>8,.2f}')
print('   (that is the accounting fees line she corrected by hand in the export)')
print()
print('--- margins and departments, August ---')
for lab in ['    Gross profit', '    Gross profit %', '    EBITDA', '    EBIT', '    Net profit %']:
    val = pl.cell(LBL[lab], AUG).value
    print(f'   {lab.strip():<26}' + (f'{val*100:>12,.1f}%' if lab.endswith('%')
                                     else f'{val:>13,.2f}'))
mh = [r for r in range(6, 190) if pl.cell(r, 1).value == 'Profit Margin Per Department'][0]
dsum = 0
for i, d in enumerate(['Production', 'Video', 'Onsite', 'Integration', 'Consulting',
                       'General Admin']):
    rr = [r for r in range(6, mh) if pl.cell(r, 1).value == '    ' + d]
    inc = pl.cell(rr[0], AUG).value or 0
    dsum += inc
    m = pl.cell(mh + 1 + i, AUG).value
    print(f'   {d:<16} revenue {inc:>12,.2f}   margin '
          + (f'{m*100:>7,.1f}%' if isinstance(m, float) else '      -'))
chk('six departments = Total Income', dsum, pl.cell(LBL['Total Income '], AUG).value)
print()
print(f'   no cost centre: {st["B22"].value} lines, {st["B37"].value:,.2f} put into General Admin')
pc = v['PL_Check']
print('   PL_Check:', [pc.cell(5 + i, 11).value for i in range(4)], '| net', pc.cell(10, 11).value)
print()
print('THE PACK TIES TO HER EXPORT' if ok else 'SOMETHING IS OUT')
