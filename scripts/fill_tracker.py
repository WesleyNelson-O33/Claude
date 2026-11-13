"""Type the three Finance cells on the tracker for the invoices Xero has and it does not.

The tracker's own rule, from its Read Me: the department sets up every job
line, Finance never types a job number, and Finance types three things off
the Xero invoice, the invoice number, the invoice date and the amount
excluding GST. So this only ever writes columns G, H and I on the Finance
sheet, onto a row the department has already created. Nothing else is
touched.

A row is only filled where the match is certain. Either the Project number
on the Xero invoice equals the Job Number on an open Finance row and that is
the only open row for that job, or, for the contract invoices that carry no
Project number, the contract reference matches the row's description and the
amount agrees. Everything else is left alone and listed for you.

    scripts/fill_tracker.py <tracker in> <tracker out> <fill.json> <extra.json>
"""
import datetime
import json
import sys

import openpyxl

COL_INV, COL_DATE, COL_AMT = 7, 8, 9      # G, H, I on Finance


def main():
    if len(sys.argv) < 5:
        print(__doc__.strip())
        return 1
    src, out, fillf, extraf = sys.argv[1:5]
    fill = json.load(open(fillf))
    fill.update(json.load(open(extraf)))

    wb = openpyxl.load_workbook(src)
    ws = wb['Finance']
    if str(ws.cell(6, COL_INV).value).strip() != 'Xero Invoice No':
        print('STOPPED: Finance column G is %r, not Xero Invoice No'
              % ws.cell(6, COL_INV).value)
        return 1
    if str(ws.cell(6, COL_AMT).value).strip() != 'Xero Invoiced Ex GST':
        print('STOPPED: Finance column I is %r' % ws.cell(6, COL_AMT).value)
        return 1

    done = []
    for inv, (row, date, net) in sorted(fill.items(), key=lambda x: x[1][0]):
        here = ws.cell(row, COL_INV).value
        if here not in (None, ''):
            print('STOPPED: Finance r%d already holds %r, not writing %s'
                  % (row, here, inv))
            return 1
        ws.cell(row, COL_INV).value = inv
        ws.cell(row, COL_DATE).value = datetime.datetime.strptime(date, '%Y-%m-%d')
        ws.cell(row, COL_DATE).number_format = 'dd/mm/yyyy'
        ws.cell(row, COL_AMT).value = net
        done.append((row, inv, date, net))

    wb.save(out)
    print('  %d Finance rows filled, columns G, H and I only' % len(done))
    for row, inv, date, net in done:
        print('     r%-4d %-11s %s %12s' % (row, inv, date, '{:,.2f}'.format(net)))
    print('  saved %s' % out)
    return 0


if __name__ == '__main__':
    sys.exit(main())
