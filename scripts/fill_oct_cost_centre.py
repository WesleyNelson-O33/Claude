"""Fill the October P&L by Cost Centre tab.

Built the way the August tab says it was built: every approved October invoice
line in Xero, by account, allocated to the cost centre of its Finance row on
the tracker. Nine invoices sit on Finance rows where the department never set a
cost centre, so they go in the Reallocation column, which is what that column
is for.

Column K is the October Xero P&L by account, read from Xero on 13 November
2026. The Check column then shows what is left per account. What is left is
credit notes, the WIP journal and deferral releases, none of which are invoice
lines. The sheet says so under the table.

    scripts/fill_oct_cost_centre.py <tracker in> <tracker out> <split.json> <cc.json>
"""
import json
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

TAB = 'Oct 26 P&L by Cost Centre'
CENTRES = ['ONSITE', 'PRODUCTION', 'VIDEO', 'INTEGRATION', 'CONSULTING', 'CTS', 'REALLOC']
COL = dict(zip(CENTRES, range(3, 10)))
XERO_PL = {
    '41100': 149149.44, '41150': 820.00, '41175': 36175.12, '42100': 231440.84,
    '42150': 29643.75, '42200': 32600.70, '42300': 5098.00, '42500': 375289.10,
    '42600': 132231.00, '42700': 16800.00, '42800': 4860.00, '42900': 33695.00,
    '45015': 5350.00, '45020': 0.00, '45025': 0.00, '47015': 3105.00,
}
FOOT = ('Invoice lines and the two October credit notes. CN-10695 and CN-10696 are '
        '12,388.80 each on 41100 ONSITE. INV-10683 sits on a Production row and a Video '
        'row, so its video line goes to VIDEO first and the balance VIDEO needs comes off '
        'Equipment Hires, which is how August did it. What the Check column still shows is '
        'the WIP journal, which is not on this tab. The October deferral release of '
        '7,873.00 on 41175 INTEGRATION is included, the way August included it. The '
        'Reallocation column is nine invoices whose Finance row has no cost centre: '
        'INV-10704, 10709, 10726, 10728, 10731, 10733, 10734, 10739 and 10740.')


def main():
    if len(sys.argv) < 5:
        print(__doc__.strip())
        return 1
    split = json.load(open(sys.argv[3]))
    cc = json.load(open(sys.argv[4]))
    wb = openpyxl.load_workbook(sys.argv[1])
    if TAB not in wb.sheetnames:
        print('STOPPED: no %r tab' % TAB)
        return 1
    ws = wb[TAB]

    # rebuild the account by cost centre grid from the invoice splits
    grid = {}
    for key, amt in split.items():
        code, _ = key.split('|')
        grid.setdefault(code, {})
    grid = {}
    raw = json.load(open(sys.argv[3]))
    for key, amt in raw.items():
        code, centre = key.split('|')
        grid.setdefault(code, {})
        grid[code][centre] = round(grid[code].get(centre, 0.0) + amt, 2)

    written = 0
    for r in range(5, 21):
        code = str(ws.cell(r, 1).value or '').strip()
        if not code:
            continue
        for centre in CENTRES:
            v = grid.get(code, {}).get(centre)
            if v:
                ws.cell(r, COL[centre]).value = round(v, 2)
                written += 1
        if code in XERO_PL:
            ws.cell(r, 11).value = XERO_PL[code]
        else:
            print('STOPPED: no Xero P&L figure for account %s' % code)
            return 1
    seen = set(str(ws.cell(r, 1).value or '').strip() for r in range(5, 21))
    stray = [c for c in grid if c not in seen]
    if stray:
        print('STOPPED: these accounts have invoice lines but no row on the tab: %s' % stray)
        return 1

    f = ws.cell(23, 1)
    f.value = FOOT
    f.font = Font(name='Arial', italic=True, size=9, color='242C39')
    f.fill = PatternFill('solid', fgColor='FFF2CC')
    f.alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells(start_row=23, start_column=1, end_row=23, end_column=12)
    ws.row_dimensions[23].height = 56
    ws['A2'] = ('October 2026 trading income by cost centre. Every approved October invoice '
                'line in Xero, by account, allocated to the cost centre of its Finance row '
                'on this tracker. Column K is the Xero P&L for October, read on 13-Nov-26. '
                'Prepared 13-Nov-26.')
    ws['A2'].fill = PatternFill('solid', fgColor='FFFFFF')
    wb.save(sys.argv[2])
    print('  %d cost centre cells written, 16 Xero P&L figures' % written)
    print('  saved %s' % sys.argv[2])
    return 0


if __name__ == '__main__':
    sys.exit(main())
