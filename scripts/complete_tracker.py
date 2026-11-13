"""Create the missing department rows on the tracker, then type the Finance cells.

Danica asked for the department section to be completed, having been told the
tracker's own rule is that the department sets up every job line. So this does
it on her instruction, and it only ever writes what Xero actually says.

Each new row copies how the same recurring line was set up last month. The CBA
chargebacks follow the September rows: job 1668, client CBA, cost centre
ONSITE, invoice type Chargeback, GST 10%, and the Xero reference as the
description. The Aware Super maintenance follows its October row, job 26022301.
The Production rows take the Project number printed on the Xero invoice.

One thing is not from a source and is flagged on every row it affects. Expected
Revenue Ex GST is the department's own expectation and Xero does not hold it.
These rows are built from the invoice, so Expected is set to the invoiced
figure. That keeps the revenue totals right but it means the variance check on
those rows is not independent. Each row says so in its Notes.

    scripts/complete_tracker.py <tracker in> <tracker out>
"""
import datetime
import sys

import openpyxl

NOTE = ("Row created by Finance 13/11/26 from the Xero invoice, the department had "
        "not set it up. Expected = invoiced, so no independent variance on this row.")

# Onsite: job, client, description, cost centre, invoice type, tax, expected, invoice, date
ONSITE = [
 ("1668","CBA","85720001CBPSCSJUN27 AFTERHOURS","ONSITE","Chargeback","GST 10%",1950.00,"INV-10712","2026-10-30"),
 ("1668","CBA","85720001CBPSCSJUN27 DEDICATED TECH","ONSITE","Chargeback","GST 10%",425.00,"INV-10714","2026-10-30"),
 ("1668","CBA","85720001CBPSCSJUN27 DEDICATED TECH","ONSITE","Chargeback","GST 10%",340.00,"INV-10717","2026-10-30"),
 ("1668","CBA","85720001CBPSCSJUN27 DEDICATED TECH","ONSITE","Chargeback","GST 10%",425.00,"INV-10718","2026-10-30"),
 ("1668","CBA","85720001CBPSCSJUN27 DEDICATED TECH","ONSITE","Chargeback","GST 10%",340.00,"INV-10719","2026-10-30"),
 ("1668","CBA","85720001MEL435CSJUN27 AFTERHOURS","ONSITE","Chargeback","GST 10%",680.00,"INV-10711","2026-10-30"),
 ("1668","CBA","85720001MEL435CSJUN27 DEDICATED TECH","ONSITE","Chargeback","GST 10%",680.00,"INV-10713","2026-10-30"),
 ("1668","CBA","85720001MEL435CSJUN27 DEDICATED TECH","ONSITE","Chargeback","GST 10%",340.00,"INV-10715","2026-10-30"),
 ("1668","CBA","85720001MEL435CSJUN27 DEDICATED TECH","ONSITE","Chargeback","GST 10%",680.00,"INV-10716","2026-10-30"),
 ("26022301","Aware Super","Aware Super Annual National Maintenance Proposal","ONSITE","Contract","GST 10%",11000.00,"INV-10652","2026-09-30"),
]
PRODUCTION = [
 ("26092105","CBA","CBA - Live Event Support October 2026 (PO3323120)","PRODUCTION","Project","GST 10%",9927.50,"INV-10727","2026-10-30"),
 ("260806","CBA","CBA - Global Markets Conference 19-21.10.2026 (PO3321201)","PRODUCTION","Project","GST 10%",4775.00,"INV-10741","2026-10-31"),
 ("26091002","Global Philanthropic","Talking Philanthropy Event B-Roll","PRODUCTION","Project","GST 10%",1249.00,"INV-10669","2026-09-30"),
 ("2609402","MUFG","SUN Investor Day 29.10.26 - AV","PRODUCTION","Project","GST 10%",30317.75,"INV-10729","2026-10-30"),
 ("26061606","MUFG","Endeavour AGM 30.10.26 - Webcast & Cameras","PRODUCTION","Project","GST 10%",7649.55,"INV-10732","2026-10-30"),
 ("26061606","MUFG","EDV Comms 30.10.2026","PRODUCTION","Project","GST 10%",2000.00,"INV-10730","2026-10-30"),
]
# invoices that go onto a Finance row that already exists and is open
DIRECT = {
 "INV-10689": (253, "2026-10-30", 7150.00),      # job 2604901, PwC Brisbane pool, exact
 "INV-10701": (250, "2026-10-30", 11907.00),     # job 25101301B, PwC Sydney pool, exact
 "INV-10731": (425, "2026-10-30", 18796.30),     # job 26061606, MUFG Endeavour AGM
}
# the half set up Onsite row that only lacks a job number
FINISH_ONSITE = (80, "1350", "INV-10685", "2026-10-30", 7434.50, 232)

COLS = dict(job=3, client=6, desc=7, cc=8, typ=9, tax=10, exp=15)
NOTES_COL = dict(Onsite=23, Production=46)
RID_COL = dict(Onsite=24, Production=47)


def finance_row_for(f, rowid):
    for r in range(7, f.max_row + 1):
        if str(f.cell(r, 1).value or '').strip() == rowid:
            return r
    return None


def main():
    if len(sys.argv) < 3:
        print(__doc__.strip())
        return 1
    wb = openpyxl.load_workbook(sys.argv[1])
    f = wb['Finance']
    written = []

    def fill_fin(row, inv, date, net):
        if f.cell(row, 7).value not in (None, ''):
            raise SystemExit('STOPPED: Finance r%d already holds %r' % (row, f.cell(row, 7).value))
        f.cell(row, 7).value = inv
        f.cell(row, 8).value = datetime.datetime.strptime(date, '%Y-%m-%d')
        f.cell(row, 8).number_format = 'dd/mm/yyyy'
        f.cell(row, 9).value = net

    for inv, (row, date, net) in sorted(DIRECT.items()):
        fill_fin(row, inv, date, net)
        written.append(('Finance only', '', row, inv, net))

    r, job, inv, date, net, frow = FINISH_ONSITE
    o = wb['Onsite']
    if o.cell(r, COLS['job']).value not in (None, ''):
        print('STOPPED: Onsite r%d already has a job number' % r)
        return 1
    o.cell(r, COLS['job']).value = job
    o.cell(r, NOTES_COL['Onsite']).value = ('Job number added by Finance 13/11/26 from the '
                                            'September row for the same contract.')
    fill_fin(frow, inv, date, net)
    written.append(('Onsite r%d completed' % r, job, frow, inv, net))

    for sheet, block in (('Onsite', ONSITE), ('Production', PRODUCTION)):
        ws = wb[sheet]
        nxt = 0
        for rr in range(5, ws.max_row + 1):
            if ws.cell(rr, COLS['job']).value not in (None, ''):
                nxt = rr
        nxt += 1
        for job, client, desc, cc, typ, tax, exp, inv, date in block:
            if ws.cell(nxt, COLS['job']).value not in (None, ''):
                print('STOPPED: %s r%d is not empty' % (sheet, nxt))
                return 1
            ws.cell(nxt, COLS['job']).value = job
            ws.cell(nxt, COLS['client']).value = client
            ws.cell(nxt, COLS['desc']).value = desc
            ws.cell(nxt, COLS['cc']).value = cc
            ws.cell(nxt, COLS['typ']).value = typ
            ws.cell(nxt, COLS['tax']).value = tax
            ws.cell(nxt, COLS['exp']).value = exp
            ws.cell(nxt, NOTES_COL[sheet]).value = NOTE
            rid = str(ws.cell(nxt, RID_COL[sheet]).value or '').strip()
            frow = finance_row_for(f, rid)
            if not frow:
                print('STOPPED: no Finance row carries Row ID %r' % rid)
                return 1
            fill_fin(frow, inv, date, exp)
            written.append(('%s r%d new (%s)' % (sheet, nxt, rid), job, frow, inv, exp))
            nxt += 1

    wb.save(sys.argv[2])
    print('  %d invoices placed' % len(written))
    for what, job, frow, inv, net in written:
        print('     %-28s job %-10s Finance r%-4d %-11s %12s'
              % (what, job or '-', frow, inv, '{:,.2f}'.format(net)))
    print('  saved %s' % sys.argv[2])
    return 0


if __name__ == '__main__':
    sys.exit(main())
