"""Check every figure the workbook computes against the same figure
worked out in Python straight from the export."""
import sys, datetime, openpyxl
sys.path.insert(0, 'recon')
from build_licence_margins import (read, signed, classify, customer, term,
                                   INC, EXP)

rows = read(sys.argv[2])
releases, deferrals = classify(rows)
# repeat the tagging main() does, so pairing matches the workbook
table = []
for desc, items in releases.items():
    items.sort(key=lambda x: x['date'])
    inc = sum(x['inc'] for x in items)
    table.append(dict(desc=desc, items=items,
                      pct=(inc + sum(x['cost'] for x in items)) / inc))
table.sort(key=lambda t: t['pct'])
for i, t in enumerate(table, 1):
    t['id'] = f'L{i:02d}'
    for rel in t['items']:
        for x in rel['lines']:
            x['pairing'], x['licence'] = 'Release', t['id']

wb = openpyxl.load_workbook(sys.argv[1], data_only=True)
bad = []

def chk(label, got, want, tol=0.005):
    ok = got is not None and abs(got - want) <= tol
    if not ok:
        bad.append(label)
    print(f'  {"OK " if ok else "BAD"}  {label:<44} '
          f'{"-" if got is None else f"{got:,.2f}":>14}  vs {want:>14,.2f}')

ws = wb['By Month']
tr = ws.max_row
ti = sum(signed(x) for x in rows if x['code'] == INC)
tc = sum(signed(x) for x in rows if x['code'] == EXP)
chk('By Month total income', ws.cell(tr, 2).value, ti)
chk('By Month total cost', ws.cell(tr, 3).value, tc)

ws = wb['Licence Margins']
lk = {t['desc']: t for t in table}
nrow = nd = 0
for r in range(6, ws.max_row + 1):
    d = ws.cell(r, 2).value
    if d not in lk:
        continue
    it = lk[d]['items']
    n = len(it)
    inc = sum(x['inc'] for x in it)
    cost = sum(x['cost'] for x in it)
    g = lambda c: ws.cell(r, c).value
    for lab, got, want in (('months', g(7), n), ('income', g(14), inc),
                           ('cost', g(15), cost),
                           ('income/mth', g(10), inc / n),
                           ('margin/mth', g(12), (inc + cost) / n),
                           ('margin %', g(13), (inc + cost) / inc)):
        if got is None or abs(got - want) > 0.01:
            bad.append(f'row {r} {lab}')
            print(f'  BAD  row {r} {d[:30]} {lab}: {got} vs {want}')
    fd = g(8).date() if isinstance(g(8), datetime.datetime) else g(8)
    ld = g(9).date() if isinstance(g(9), datetime.datetime) else g(9)
    if fd != it[0]['date'] or ld != it[-1]['date']:
        bad.append(f'row {r} dates')
        print(f'  BAD  row {r} dates {fd}/{ld} vs '
              f'{it[0]["date"]}/{it[-1]["date"]}')
    else:
        nd += 1
    nrow += 1
print(f'  OK   Licence Margins: {nrow} rows checked, {nd} with both dates '
      f'matching')

ws = wb['Read This First']
for label, lo, hi in (('FY2025', (2024, 7), (2025, 6)),
                      ('FY2026', (2025, 7), (2026, 6)),
                      ('FY2027 to 30 Sep 2026', (2026, 7), (2026, 9))):
    sel = [x for x in rows if lo <= (x['date'].year, x['date'].month) <= hi]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == label:
            chk(f'{label} income', ws.cell(r, 2).value,
                sum(signed(x) for x in sel if x['code'] == INC))
            chk(f'{label} cost', ws.cell(r, 3).value,
                sum(signed(x) for x in sel if x['code'] == EXP))

sel27 = [x for x in rows
         if (2026, 7) <= (x['date'].year, x['date'].month) <= (2026, 9)]
for label, pair in (('Licences released by journal', 'Release'),
                    ('Deferred out to WIP', 'Deferral'),
                    ('Cost and income with no release journal', 'Unpaired')):
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == label:
            chk(f'FY27 {pair} cost', ws.cell(r, 3).value,
                sum(signed(x) for x in sel27
                    if x['pairing'] == pair and x['code'] == EXP))

ws = wb['Licence by Month']
gt = sum(ws.cell(r, ws.max_column).value or 0
         for r in range(5, ws.max_row + 1))
chk('Licence by Month, all margins',
    gt, sum(x['inc'] + x['cost'] for t in table for x in t['items']))

ws = wb['Unpaired Cost and Income']
chk('Unpaired total', ws.cell(ws.max_row, 7).value,
    sum(signed(x) for x in rows if x['pairing'] == 'Unpaired'))

ws = wb['Deferrals to WIP']
chk('Deferrals income', ws.cell(ws.max_row, 3).value,
    sum(d['inc'] for d in deferrals))
chk('Deferrals cost', ws.cell(ws.max_row, 4).value,
    sum(d['cost'] for d in deferrals))

print()
print('  FAILURES:', len(bad) if bad else 'none')
sys.exit(1 if bad else 0)
