"""The two payroll tax journals for a month, built the way October was.

This is build_oct_journals.py with the month, the date, the NSW figure and the
accrual pots lifted out as arguments, so a new copy of the script is not needed
every month. Run against October it writes the October journals that were
posted, line for line, which is the test that the method has not moved.

The method:

  NSW            the month's NSW payroll tax, from the Main & NSW tab, spread
                 across the NSW jobs in proportion to units.
  other states   each state's own monthly accrual spread across that state's
                 jobs in proportion to units. The accruals come off the
                 Other States = Accrual tab, which holds last year's actual
                 divided by twelve and rounded to the nearest hundred.
                 ACT is nil, dropped from August 2026.

  A line that rounds to nil is left off. The rounding difference goes on the
  largest line so the journal ties exactly to a round pot.

Cost centres are taken from the previous month's journals where the job
appears in them, so the coding carries over rather than being invented. A job
that is new this month has its cost centre read off the tag at the end of its
name, and every one of those is listed at the end so it can be checked.

The state on each row is the one the split file carries. Where the Employment
Hero payroll tax report puts an employee in a different state, the report wins,
because it groups by residential address and that is the basis the return is
lodged on. Those employees are named in MOVED below, with the month they
apply to, so the override is on the page rather than hidden in the data.

    ptax/build_journals.py <split file> <prior NSW journal> <prior other
        states journal> <month name> <journal date> <NSW tax> <out dir>

Example:

    ptax/build_journals.py split.xlsx ptax/2026-10_..._NSW_....csv \\
        ptax/2026-10_..._ACCRUAL.csv November 30/11/2026 6768.53 ptax
"""
import csv
import os
import re
import sys
from collections import defaultdict

import openpyxl

LIABILITY, EXPENSE = '21440', '65106'

# Other States = Accrual tab, the monthly accrual per state. ACT was dropped
# from August 2026 and the tab carries the note saying so.
POTS = {'QLD': 500.0, 'SA': 200.0, 'VIC': 1300.0, 'WA': 700.0, 'ACT': 0.0}

# month -> employee -> the state the Employment Hero payroll tax report puts
# them in. The split file's State column disagrees for these people.
MOVED = {
    # November 2026: the split file has Nicol Schildknecht in New South Wales
    # and the report has them in Victoria, 462.05 of wages and 55.45 of super.
    # Both their November jobs are Melbourne work.
    'November': {'Nicol Schildknecht': 'VIC'},
}

HDR = ['*Narration', '*Date', 'Description', '*AccountCode', '*TaxRate',
       '*Amount', 'TrackingName1', 'TrackingOption1', 'TrackingName2',
       'TrackingOption2']

TAGS = {'ONS': 'Onsite', 'PRD': 'Production', 'VID': 'Video',
        'INT': 'Integration', 'CONS': 'Consulting', 'CTS': 'CTS'}


def units_by_month(path):
    """Units per state per job, from the split file, with the employees the
    payroll tax report puts in another state moved to that state."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    it = wb['EH data (2)'].iter_rows(values_only=True)
    next(it)                      # the totals strip above the header
    next(it)                      # the header
    out = defaultdict(lambda: defaultdict(float))
    moved = defaultdict(float)
    for r in it:
        if all(v is None for v in r) or r[10] is None:
            continue
        month = str(r[0]).strip()
        state = str(r[21]).strip()
        who = str(r[6]).strip()
        to = MOVED.get(month, {}).get(who)
        if to and to != state:
            moved[(month, who, state, to)] += float(r[10])
            state = to
        job = f'{r[13]} - {r[14]}' if r[13] is not None else str(r[14])
        out[month][(state, job)] += float(r[10])
    wb.close()
    for (month, who, frm, to), units in sorted(moved.items()):
        print(f'  {who} moved from {frm} to {to} for {month}, '
              f'{units:,.2f} units, per the payroll tax report')
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
    """The cost centre as the previous month's journals coded this job, or,
    for a job that is new this month, from the tag at the end of its name. A
    tag can carry extra words, as in [ONS non-chargeable], so only the code at
    the front of it is read."""
    if job in known:
        return known[job], False
    m = re.search(r'\[\s*([A-Za-z]+)[^\]]*\]\s*$', job)
    if m and m.group(1).upper() in TAGS:
        return TAGS[m.group(1).upper()], True
    # some jobs carry no tag at all. They were all coded to Onsite, and they
    # all name a site rather than a department.
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
    biggest = None
    if plug:
        biggest = max(share, key=lambda j: share[j])
        share[biggest] = round(share[biggest] + plug, 2)
    return share, plug, biggest


def amount(x):
    """The way the journals are written: no trailing zeros on a round figure."""
    t = f'{x:.2f}'.rstrip('0').rstrip('.')
    return t if t not in ('', '-0') else '0'


def write(path, narration, date, credit_desc, lines, pot):
    with open(path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(HDR)
        w.writerow([narration, date, credit_desc, LIABILITY, 'BAS Excluded',
                    amount(-pot), 'Cost Centres', 'CTS', 'Job Numbers',
                    '9000 - OFFICE / ADMIN [CTS]'])
        for job, amt, centre in lines:
            w.writerow([narration, date, job, EXPENSE, 'BAS Excluded',
                        amount(amt), 'Cost Centres', centre, 'Job Numbers',
                        job])


def main():
    if len(sys.argv) != 8:
        print(__doc__)
        return 1
    split, prior_nsw, prior_oth, month, date, nsw_tax, outdir = sys.argv[1:]
    nsw_tax = float(nsw_tax)
    year = date.split('/')[-1]
    stamp = '%s-%02d' % (year, int(date.split('/')[1]))

    out_nsw = os.path.join(outdir, f'{stamp}_Payroll_Tax_NSW_Xero_Journal.csv')
    out_oth = os.path.join(
        outdir, f'{stamp}_Payroll_Tax_Alloc_Other_States_'
                f'Xero_Journal_-_ACCRUAL.csv')

    U = units_by_month(split)
    nsw_centres, nsw_credit = read_journal(prior_nsw)
    oth_centres, oth_credit = read_journal(prior_oth)
    known = {**nsw_centres, **oth_centres}
    units = U[month]
    if not units:
        print(f'  STOPPED, no {month} rows in the split file')
        return 1

    guessed, report = [], []

    # ---------------------------------------------------------- NSW
    jobs = {j: u for (st, j), u in units.items() if st == 'NSW'}
    share, plug, on = allocate(jobs, nsw_tax)
    lines = []
    for job in sorted(share, key=lambda j: -share[j]):
        centre, derived = centre_for(job, known)
        if centre is None:
            print(f'  STOPPED, no cost centre for {job}')
            return 1
        if derived:
            guessed.append((job, centre))
        lines.append((job, share[job], centre))
    write(out_nsw, f'Payroll Tax NSW {month} {year}', date,
          nsw_credit['Description'], lines, nsw_tax)
    report.append(('NSW', len(jobs), len(lines), nsw_tax, plug, on))

    # ------------------------------------------------- the other states
    merged, plugs = defaultdict(float), []
    for st, pot in POTS.items():
        sj = {j: u for (s, j), u in units.items() if s == st}
        if pot == 0:
            if sj:
                print(f'  note: {st} has {len(sj)} jobs with units in '
                      f'{month} but its accrual is nil')
            continue
        if not sj:
            print(f'  STOPPED, {st} has a pot of {pot} but no {month} units')
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
    write(out_oth, f'Payroll Tax ACT/QLD/SA/VIC/WA {month} {year} Accrual',
          date, oth_credit['Description'], lines, pot_total)
    report.append(('other states', len(merged), len(lines), pot_total,
                   None, None))

    # ---------------------------------------------------------- checks
    for path, name in ((out_nsw, 'NSW'), (out_oth, 'other states')):
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
        print(f'\n  {len(guessed)} job(s) are not in the previous month\'s '
              f'journals, so the cost centre is read off the tag. Please check:')
        for j, c in sorted(set(guessed)):
            print(f'    {c:<12} {j}')
    else:
        print('\n  every job appears in the previous month\'s journals, so '
              'every cost centre carries over')
    return 0


if __name__ == '__main__':
    sys.exit(main())
