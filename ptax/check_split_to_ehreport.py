"""The monthly control: the split file against the Employment Hero
residential address payroll tax report.

    split gross + SG super + employer contribution
      less any ad hoc pay run
      = the report's grand total, state by state

The report leaves ad hoc pay runs out. That is the whole of the difference
in October 2026 and it is worth 196.72 of New South Wales payroll tax, so
the control has to be run before the return is lodged, not after.

September 2026 reconciles on the same basis once the exempt per diems are
taken out, 300.00 on New South Wales and 525.00 on South Australia. The
October report has no exempt per diems at all.

Run with the split file, then the month, then the report's grand total per
state as STATE=amount pairs, for example:

    check_split_to_ehreport.py split.xlsx October \\
        NSW=187262.28 QLD=14489.10 SA=6045.37 VIC=19799.31 WA=8505.86

The split file must have been recalculated, because the state comes from a
lookup.
"""
import sys
import datetime
from collections import defaultdict

import openpyxl

# the employer contribution the payroll report carries outside SG super
EMPLOYER_CONTRIBUTION = {'NSW': 2001.04}
STATES = ('NSW', 'QLD', 'SA', 'VIC', 'WA', 'ACT')


def main():
    if len(sys.argv) < 4:
        print(__doc__)
        return 1
    path, month = sys.argv[1], sys.argv[2]
    report = {}
    for pair in sys.argv[3:]:
        st, _, amt = pair.partition('=')
        report[st.strip().upper()] = float(amt)

    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    agg = defaultdict(lambda: defaultdict(float))
    runs = defaultdict(lambda: defaultdict(float))
    it = wb['EH data (2)'].iter_rows(values_only=True)
    next(it)
    next(it)
    for r in it:
        if all(v is None for v in r) or r[4] is None:
            continue
        if str(r[0]).strip() != month:
            continue
        st = str(r[21]).strip() if r[21] else 'NO STATE'
        g, s = float(r[18] or 0), float(r[20] or 0)
        pay = r[2].date() if r[2] else None
        agg[st]['gross'] += g
        agg[st]['sg'] += s
        runs[pay]['gross'] += g
        runs[pay]['sg'] += s
        runs[pay]['rows'] += 1
    wb.close()

    if 'NO STATE' in agg:
        print(f'  STOP. {int(agg["NO STATE"]["gross"]):,} of gross has no '
              f'state. The lookup has not been filled down, or the file has '
              f'not been recalculated.')
        return 1

    print(f'  pay runs in {month}')
    main_runs = sorted(runs, key=str)
    for d in main_runs:
        e = runs[d]
        print(f'    paid {d}  {int(e["rows"]):>5} rows  '
              f'gross {e["gross"]:>12,.2f}  super {e["sg"]:>10,.2f}')
    if len(main_runs) > 2:
        print(f'    {len(main_runs)} pay dates. More than two means an ad hoc '
              f'run, and the report will not have it.')

    print(f'\n  {"":<6}{"split total":>15}{"the report":>14}'
          f'{"difference":>14}')
    tot_split = tot_rep = 0.0
    for st in STATES:
        if st not in agg and st not in report:
            continue
        split = round(agg[st]['gross'] + agg[st]['sg']
                      + EMPLOYER_CONTRIBUTION.get(st, 0.0), 2)
        rep = report.get(st, 0.0)
        tot_split += split
        tot_rep += rep
        flag = '' if abs(split - rep) < 0.02 else '   <-- explain this'
        print(f'  {st:<6}{split:>15,.2f}{rep:>14,.2f}'
              f'{split - rep:>+14,.2f}{flag}')
    gap = round(tot_split - tot_rep, 2)
    print(f'  {"total":<6}{tot_split:>15,.2f}{tot_rep:>14,.2f}{gap:>+14,.2f}')

    print()
    if abs(gap) < 0.02:
        print('  The two agree. Nothing to explain.')
    else:
        print(f'  They differ by {gap:,.2f}. Every dollar of that has to be '
              f'named before the return goes in. The usual causes are an ad '
              f'hoc pay run, which the report leaves out, and exempt per '
              f'diems, which it takes out.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
