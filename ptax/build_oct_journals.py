"""The two October 2026 payroll tax journals, built the way September was.

September is the model and the method reproduces it. Taking the September
rows out of the Employment Hero split file and spreading each pot across the
jobs by units gives, line for line, the two journals that were posted. The
only differences are the rounding plugs: the other states journal was one
cent off and the NSW journal two cents, and both were plugged onto a single
line so the journal ties to a round pot.

The method:

  NSW            the month's NSW payroll tax, spread across the NSW jobs in
                 proportion to units.
  other states   each state's own monthly accrual spread across that state's
                 jobs in proportion to units. QLD 500, SA 200, VIC 1,300,
                 WA 700. ACT is nil, dropped from August 2026.

  A line that rounds to nil is left off. The rounding difference goes on the
  largest line so the journal ties exactly.

October inputs:

  NSW payroll tax 5,800.27, from Main & NSW cell E23 once the October
  Employment Hero report figures are in. All five states tie to that report.

Cost centres are taken from the September journals where the job appears in
them, so the coding carries over rather than being invented. A job that is
new in October has its cost centre read off the tag at the end of its name,
and every one of those is listed at the end so it can be checked.

Run with the split file, the September NSW journal and the September other
states journal.
"""
import csv
import re
import sys
from collections import defaultdict

import openpyxl

OUT_NSW = 'ptax/2026-10_Payroll_Tax_NSW_Xero_Journal.csv'
OUT_OTH = ('ptax/2026-10_Payroll_Tax_Alloc_Other_States_'
           'Xero_Journal_-_ACCRUAL.csv')

MONTH = 'October'
DATE = '31/10/2026'
LIABILITY, EXPENSE = '21440', '65106'
NSW_TAX = 5800.27
POTS = {'QLD': 500.0, 'SA': 200.0, 'VIC': 1300.0, 'WA': 700.0, 'ACT': 0.0}

HDR = ['*Narration', '*Date', 'Description', '*AccountCode', '*TaxRate',
       '*Amount', 'TrackingName1', 'TrackingOption1', 'TrackingName2',
       'TrackingOption2']

TAGS = {'ONS': 'Onsite', 'PRD': 'Production', 'VID': 'Video',
        'INT': 'Integration', 'CONS': 'Consulting', 'CTS': 'CTS'}


def units_by_month(path):
    """Units per state per job, from the split file."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    it = wb['EH data (2)'].iter_rows(values_only=True)
    next(it)                      # the totals strip above the header
    next(it)                      # the header
    out = defaultdict(lambda: defaultdict(float))
    for r in it:
        if all(v is None for v in r) or r[10] is None:
            continue
        job = f'{r[13]} - {r[14]}' if r[13] is not None else str(r[14])
        out[str(r[0]).strip()][(str(r[21]).strip(), job)] += float(r[10])
    wb.close()
    return out


def read_journal(path):
    rows = list(csv.DictReader(open(path, newline='', encoding='utf-8-sig')))
    centres, credit = {}, None
    for r in rows:
        if r['*AccountCode'] == EXPENSE:
            centres[r['TrackingOption2']] = r['TrackingOption1']
        else:
            credit = r
    return centres, credit


def centre_for(job, known):
    """The cost centre as the September journals coded this job, or, for a
    job that is new this month, from the tag at the end of its name. A tag
    can carry extra words, as in [ONS non-chargeable], so only the code at
    the front of it is read."""
    if job in known:
        return known[job], False
    m = re.search(r'\[\s*([A-Za-z]+)[^\]]*\]\s*$', job)
    if m and m.group(1).upper() in TAGS:
        return TAGS[m.group(1).upper()], True
    # some jobs carry no tag at all. September coded every one of those to
    # Onsite, and they all name a site rather than a department.
    if re.match(r'^\d+\s*-\s*ONS\b', job) or 'ONSITE' in job.upper():
        return 'Onsite', True
    return None, True


def allocate(jobs, pot):
    """Spread a pot across jobs by units, drop the nil lines, and put the
    rounding difference on the largest line so it ties exactly."""
    total = sum(jobs.values())
    share = {j: round(pot * u / total, 2) for j, u in jobs.items()}
    share = {j: v for j, v in share.items() if v != 0}
    plug = round(pot - sum(share.values()), 2)
    if plug:
        biggest = max(share, key=lambda j: share[j])
        share[biggest] = round(share[biggest] + plug, 2)
    return share, plug, (biggest if plug else None)


def amount(x):
    """The way September wrote them: no trailing zeros on a round figure."""
    t = f'{x:.2f}'.rstrip('0').rstrip('.')
    return t if t not in ('', '-0') else '0'


def write(path, narration, credit_desc, lines, pot):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(HDR)
        w.writerow([narration, DATE, credit_desc, LIABILITY, 'BAS Excluded',
                    amount(-pot), 'Cost Centres', 'CTS', 'Job Numbers',
                    '9000 - OFFICE / ADMIN [CTS]'])
        for job, amt, centre in lines:
            w.writerow([narration, DATE, job, EXPENSE, 'BAS Excluded',
                        amount(amt), 'Cost Centres', centre, 'Job Numbers',
                        job])


def main():
    if len(sys.argv) != 4:
        print(__doc__)
        return 1
    U = units_by_month(sys.argv[1])
    nsw_centres, nsw_credit = read_journal(sys.argv[2])
    oth_centres, oth_credit = read_journal(sys.argv[3])
    known = {**nsw_centres, **oth_centres}
    oct_units = U[MONTH]
    if not oct_units:
        print(f'  no {MONTH} rows in the split file')
        return 1

    guessed = []
    report = []

    # ---------------------------------------------------------- NSW
    jobs = {j: u for (st, j), u in oct_units.items() if st == 'NSW'}
    share, plug, on = allocate(jobs, NSW_TAX)
    lines = []
    for job in sorted(share, key=lambda j: -share[j]):
        centre, derived = centre_for(job, known)
        if centre is None:
            print(f'  STOPPED, no cost centre for {job}')
            return 1
        if derived:
            guessed.append((job, centre))
        lines.append((job, share[job], centre))
    write(OUT_NSW, f'Payroll Tax NSW {MONTH} 2026',
          nsw_credit['Description'], lines, NSW_TAX)
    report.append(('NSW', len(jobs), len(lines), NSW_TAX, plug, on))

    # ------------------------------------------------- the other states
    merged, plugs = defaultdict(float), []
    for st, pot in POTS.items():
        sj = {j: u for (s, j), u in oct_units.items() if s == st}
        if pot == 0:
            if sj:
                print(f'  note: {st} has {len(sj)} jobs with units in '
                      f'{MONTH} but its accrual is nil')
            continue
        if not sj:
            print(f'  STOPPED, {st} has a pot of {pot} but no {MONTH} units')
            return 1
        share, plug, on = allocate(sj, pot)
        plugs.append((st, plug, on))
        for j, v in share.items():
            merged[j] = round(merged[j] + v, 2)
    lines = []
    for job in sorted(merged, key=lambda j: -merged[j]):
        centre, derived = centre_for(job, known)
        if centre is None:
            print(f'  STOPPED, no cost centre for {job}')
            return 1
        if derived:
            guessed.append((job, centre))
        lines.append((job, merged[job], centre))
    pot_total = sum(POTS.values())
    write(OUT_OTH,
          f'Payroll Tax ACT/QLD/SA/VIC/WA {MONTH} 2026 Accrual',
          oth_credit['Description'], lines, pot_total)
    report.append(('other states', len(merged), len(lines), pot_total,
                   None, None))

    # ---------------------------------------------------------- checks
    for path, name in ((OUT_NSW, 'NSW'), (OUT_OTH, 'other states')):
        rows = list(csv.DictReader(open(path, newline='', encoding='utf-8-sig')))
        by = defaultdict(float)
        for r in rows:
            by[r['*AccountCode']] += float(r['*Amount'])
        out = round(sum(by.values()), 2)
        print(f'[{name}]  {len(rows)} lines, '
              f'{LIABILITY} {by[LIABILITY]:,.2f}, '
              f'{EXPENSE} {by[EXPENSE]:,.2f}, balances: {out == 0}')
        print(f'  written to {path}')
    print()
    for st, njobs, nlines, pot, plug, on in report:
        extra = '' if plug is None else \
            f', rounding plug of {plug:+.2f} on {on}'
        print(f'  {st}: {njobs} jobs, {nlines} lines, pot {pot:,.2f}{extra}')
    for st, plug, on in plugs:
        if plug:
            print(f'    {st} plug {plug:+.2f} on {on}')
    if guessed:
        print(f'\n  {len(guessed)} job(s) are not in the September journals, '
              f'so the cost centre is read off the tag. Please check:')
        for j, c in sorted(set(guessed)):
            print(f'    {c:<12} {j}')
    else:
        print('\n  every job appears in the September journals, so every '
              'cost centre carries over')
    return 0


if __name__ == '__main__':
    sys.exit(main())
