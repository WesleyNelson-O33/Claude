"""Staff expense reimbursement journals for the two fortnights to 13 November.

Copies journal 37007, the one Danica sent as the model:

  narration   EH Expense Reimbursement <run> (<start>-<end>), dates with dots
  date        the pay period ending date
  description the claim's location, which is the job name without its number
  amount      as claimed, GST inclusive, so the draft is set to Tax Inclusive
              after it is imported
  tracking    the claim's cost centre and its full job number
  credit      one line to 21410 Payroll Accruals Payable, CTS and 9000

The run numbers are read off the ledger rather than guessed: 37008 is already
posted for the fortnight 03/10/2026 to 16/10/2026, so these two are 37009 and
37010.

Account codes come from the October general ledger in the controller pack, so
they are this organisation's own codes and not a guess.

    payexp/build_nov_staff_exp.py <audit 20261030.xlsx> <audit 20261113.xlsx>
        <out dir>
"""
import csv
import os
import sys
from decimal import Decimal, ROUND_HALF_UP

import openpyxl

HEAD = ['*Narration', '*Date', 'Description', '*AccountCode', '*TaxRate',
        '*Amount', 'TrackingName1', 'TrackingOption1', 'TrackingName2',
        'TrackingOption2']

# Employment Hero expense category -> (account, tax rate). Where the account
# depends on the department, the tag at the end of the job name picks it.
ACCT = {
    'Direct Travel (chargeable)':
        ({'[PRD]': '58115', '[ONS]': '58110'}, 'GST on Expenses'),
    'Direct Freight (chargeable)':
        ({'[PRD]': '58215', '[ONS]': '58210'}, 'GST on Expenses'),
    'Taxis/Parking (non-chargeable)': ('61903', 'GST on Expenses'),
    # the category name says the claim carries no GST
    'Staff Amenities - Exc GST': ('65110', 'BAS Excluded'),
    # coded to OFFICE / ADMIN, so the overhead account rather than the cost
    # of sales one at 51500
    'Subscription & Licences': ('61400', 'GST on Expenses'),
    'Client Entertainment': ('64850', 'GST on Expenses'),
}

CC = {'[CTS]': 'CTS', '[PRD]': 'PRODUCTION', '[ONS]': 'ONSITE',
      '[VID]': 'VIDEO', '[INT]': 'INTEGRATION', '[CONS]': 'CONSULTING'}

# location -> the full job number string, as the ledger writes it
JOBS = {
    'OFFICE / ADMIN [CTS]': '9000 - OFFICE / ADMIN [CTS]',
    'APA - Qtly All Hands (BNE) - 4.11.26 [PRD]':
        '26061607 - APA - Qtly All Hands (BNE) - 4.11.26 [PRD]',
    'Ventia - CEO Town Hall - 28.10.26 [PRD]':
        '26093002 - Ventia - CEO Town Hall - 28.10.26 [PRD]',
    'Aware Super AV Maintenance [ONS]':
        '26022301 - Aware Super AV Maintenance [ONS]',
}

RUNS = {'17/10/2026 - 30/10/2026': '37009',
        '31/10/2026 - 13/11/2026': '37010'}

CREDIT = '21410'


def money(x):
    return Decimal(str(x)).quantize(Decimal('0.01'), ROUND_HALF_UP)


def build(path, outdir):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    summary = {r[0]: r[1] for r in wb['Summary'].iter_rows(values_only=True)
               if r and r[0]}
    period = str(summary['Report Period']).strip()
    if period not in RUNS:
        raise SystemExit('no run number held for the period %r' % period)
    run = RUNS[period]
    start, end = [p.strip() for p in period.split(' - ')]
    narration = 'EH Expense Reimbursement %s (%s-%s)' % (
        run, start.replace('/', '.'), end.replace('/', '.'))
    # the journal is dated the pay period ending date, the way 37007 was
    date = '%d/%d/%s' % tuple([int(x) for x in end.split('/')[:2]]
                              + [end.split('/')[2]])

    rows, total, notes = [], Decimal('0'), []
    stated = None
    for r in list(wb['Expenses'].iter_rows(values_only=True))[1:]:
        if r[0] == 'Total Expenses':
            stated = money(r[7])
            continue
        category, location, amount = r[4], str(r[5]).strip(), money(r[7])
        if category not in ACCT:
            raise SystemExit('no account held for the category %r' % category)
        if location not in JOBS:
            raise SystemExit('no job number held for %r' % location)
        tag = location[location.rindex('['):]
        accounts, tax = ACCT[category]
        account = accounts if isinstance(accounts, str) else accounts[tag]
        rows.append([narration, date, location, account, tax, str(amount),
                     'Cost Centres', CC[tag], 'Job Numbers', JOBS[location]])
        total += amount
        notes.append((category, location, amount, account, tax))
    if stated is None or stated != total:
        raise SystemExit('the claims add to %s and the report says %s'
                         % (total, stated))
    rows.append([narration, date, narration, CREDIT, 'BAS Excluded',
                 str(-total), 'Cost Centres', 'CTS', 'Job Numbers',
                 '9000 - OFFICE / ADMIN [CTS]'])

    out = os.path.join(outdir, 'Staff Expenses %s-%s Xero Journal.csv'
                       % (start.replace('/', '.'), end.replace('/', '.')))
    with open(out, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(HEAD)
        w.writerows(rows)

    print('%s  %s' % (run, period))
    print('   dated %s, %d expense lines, %s claimed' % (date, len(notes), total))
    by = {}
    for category, _loc, amount, account, tax in notes:
        key = (account, category, tax)
        by[key] = by.get(key, Decimal('0')) + amount
    for (account, category, tax), amount in sorted(by.items()):
        print('   %-6s %-32s %-16s %9s' % (account, category[:32], tax, amount))
    print('   %-6s %-32s %-16s %9s' % (CREDIT, 'Payroll Accruals Payable',
                                       'BAS Excluded', -total))
    print('   written to %s' % out)
    print()
    wb.close()


def main():
    if len(sys.argv) != 4:
        print(__doc__)
        return 1
    outdir = sys.argv[3]
    os.makedirs(outdir, exist_ok=True)
    for path in sys.argv[1:3]:
        build(path, outdir)
    return 0


if __name__ == '__main__':
    sys.exit(main())
