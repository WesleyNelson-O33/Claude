"""Set the missing cost centres on the Production sheet, and move the October
cost centre tab to match.

Ten Production rows had no cost centre. Nine of them carry an October invoice,
69,421.35 in all, which is why Month-End showed a 'No cost centre set' line.
The tenth has no invoice yet. All ten are the main job, not the V job, and the
V jobs beside three of them already say VIDEO, so all ten are PRODUCTION.

Xero already has the department on every one of these invoices, so this only
brings the tracker into line with the P&L. It changes nothing in Xero.

The October cost centre tab has to move with it, or the reconciliation breaks:
those nine amounts come out of the Reallocation column and into PRODUCTION.

    scripts/set_cost_centres.py <tracker in> <tracker out> <split.json>
"""
import json
import sys

import openpyxl

ROWS = [133, 135, 137, 139, 140, 141, 142, 143, 144, 145]
JOBS = ['26011407', '26061605', '26061611', '26093002', '2610604',
        '2610605', '2609402', '26091604', '26061606', '26061606']
CC_COL, JOB_COL, NOTE_COL = 8, 3, 46
TAB = 'Oct 26 P&L by Cost Centre'
COL = dict(ONSITE=3, PRODUCTION=4, VIDEO=5, INTEGRATION=6, CONSULTING=7, CTS=8, REALLOC=9)
NOTE = 'Cost centre set by Finance 13/11/26 to match the Xero tracking on the invoice.'


def main():
    if len(sys.argv) < 4:
        print(__doc__.strip())
        return 1
    wb = openpyxl.load_workbook(sys.argv[1])
    p = wb['Production']
    for r, job in zip(ROWS, JOBS):
        here = str(p.cell(r, JOB_COL).value or '').strip()
        if here != job:
            print('STOPPED: Production r%d is job %r, expected %r' % (r, here, job))
            return 1
        if p.cell(r, CC_COL).value not in (None, ''):
            print('STOPPED: Production r%d already has cost centre %r'
                  % (r, p.cell(r, CC_COL).value))
            return 1
        p.cell(r, CC_COL).value = 'PRODUCTION'
        old = str(p.cell(r, NOTE_COL).value or '')
        p.cell(r, NOTE_COL).value = (old + ' ' + NOTE).strip()
    print('  %d Production rows set to PRODUCTION' % len(ROWS))

    # the tab moves with it
    split = json.load(open(sys.argv[3]))
    moved = {}
    for key in list(split):
        code, centre = key.split('|')
        if centre != 'REALLOC':
            continue
        amt = split.pop(key)
        split['%s|PRODUCTION' % code] = round(split.get('%s|PRODUCTION' % code, 0) + amt, 2)
        moved[code] = amt
    ws = wb[TAB]
    for r in range(5, 21):
        code = str(ws.cell(r, 1).value or '').strip()
        if not code:
            continue
        for centre, c in COL.items():
            v = split.get('%s|%s' % (code, centre))
            ws.cell(r, c).value = round(v, 2) if v else None
    total = sum(moved.values())
    print('  %s moved out of Reallocation and into PRODUCTION on the tab' % '{:,.2f}'.format(total))
    for k, v in sorted(moved.items()):
        print('     %-6s %12s' % (k, '{:,.2f}'.format(v)))
    f = ws.cell(23, 1)
    f.value = str(f.value or '').replace(
        'The Reallocation column is nine invoices whose Finance row has no cost centre: '
        'INV-10704, 10709, 10726, 10728, 10731, 10733, 10734, 10739 and 10740.',
        'The nine invoices that had no cost centre are now PRODUCTION, so the Reallocation '
        'column is nil.')
    wb.save(sys.argv[2])
    print('  saved %s' % sys.argv[2])
    return 0


if __name__ == '__main__':
    sys.exit(main())
