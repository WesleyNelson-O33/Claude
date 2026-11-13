"""What is still missing from the tracker after the Finance cells were filled.

Each of these needs a department row before Finance can type anything, because
the tracker's rule is that Finance never types a job number.

    scripts/tracker_leftovers.py <out.xlsx>
"""
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

INK, PANEL, AMBER = '242C39', 'F4F4F4', 'FFF2CC'
MONEY = '#,##0.00'

# group, invoice, date, client, net, job on the Xero invoice, what is needed
LEFT = [
 ("Onsite contract extras, no job row",
  "INV-10712","2026-10-30","CBA Place South",1950.00,"none","CBPS AFTERHOURS. Onsite needs a row on the CBA Place South contract job."),
 ("Onsite contract extras, no job row",
  "INV-10714","2026-10-30","CBA Place South",425.00,"none","CBPS DEDICATED TECH. Onsite needs a row."),
 ("Onsite contract extras, no job row",
  "INV-10717","2026-10-30","CBA Place South",340.00,"none","CBPS DEDICATED TECH. Onsite needs a row."),
 ("Onsite contract extras, no job row",
  "INV-10718","2026-10-30","CBA Place South",425.00,"none","CBPS DEDICATED TECH. Onsite needs a row."),
 ("Onsite contract extras, no job row",
  "INV-10719","2026-10-30","CBA Place South",340.00,"none","CBPS DEDICATED TECH. Onsite needs a row."),
 ("Onsite contract extras, no job row",
  "INV-10711","2026-10-30","CBA Melbourne",680.00,"none","MEL AFTERHOURS. Onsite needs a row."),
 ("Onsite contract extras, no job row",
  "INV-10713","2026-10-30","CBA Melbourne",680.00,"none","MEL DEDICATED TECH. Onsite needs a row."),
 ("Onsite contract extras, no job row",
  "INV-10715","2026-10-30","CBA Melbourne",340.00,"none","MEL DEDICATED TECH. Onsite needs a row."),
 ("Onsite contract extras, no job row",
  "INV-10716","2026-10-30","CBA Melbourne",680.00,"none","MEL DEDICATED TECH. Onsite needs a row."),
 ("Onsite contract extras, no job row",
  "INV-10685","2026-10-30","CBA Brisbane",7434.50,"none","85720001BRISCSJUN27. The Brisbane contract row for October is not set up."),
 ("Onsite contract extras, no job row",
  "INV-10689","2026-10-30","PwC Services",7150.00,"none","PO2601702371. The PwC Brisbane pool row for October is not set up."),
 ("Onsite contract extras, no job row",
  "INV-10701","2026-10-30","PwC Services",11907.00,"none","PO2601702104. The PwC Sydney pool row for October is not set up."),
 ("Onsite contract extras, no job row",
  "INV-10652","2026-09-30","Aware Super",11000.00,"none","Annual National Maintenance, September. The October one, INV-10700, is on the tracker. September is not."),
 ("Job has no row at all",
  "INV -10585","2026-08-25","MUFG Corporate Markets",21703.75,"26061104","Production needs a row for job 26061104. Note the invoice number in Xero has a space in it."),
 ("Job has no row at all",
  "INV-10584","2026-08-25","MUFG Corporate Markets",6948.50,"26061104","Second invoice on the same job 26061104, so a second row."),
 ("Job has no row at all",
  "INV-10727","2026-10-30","CBA Place South",9927.50,"26092105","Production needs a row for job 26092105, CBA Live Event Support October."),
 ("Job has no row at all",
  "INV-10741","2026-10-31","CBA Place South",4775.00,"260806","Xero says job 260806. The tracker has 2608606 for the Global Markets Conference. One of the two numbers is wrong, please check which."),
 ("Job has no row at all",
  "INV-10669","2026-09-30","Global Philanthropic",1249.00,"26091002","Production needs a row for job 26091002, Talking Philanthropy B-Roll."),
 ("More invoices than rows",
  "INV-10729","2026-10-30","MUFG Corporate Markets",30317.75,"2609402","Job 2609402 has one row and Xero raised two invoices. INV-10728 took the row. This one needs a second row."),
 ("More invoices than rows",
  "INV-10730","2026-10-30","MUFG Corporate Markets",2000.00,"26061606","Job 26061606 has two open rows and Xero raised three invoices. I will not guess which invoice belongs on which row."),
 ("More invoices than rows",
  "INV-10731","2026-10-30","MUFG Corporate Markets",18796.30,"26061606","Same job. Tell me the pairing and I will type them."),
 ("More invoices than rows",
  "INV-10732","2026-10-30","MUFG Corporate Markets",7649.55,"26061606","Same job. Tell me the pairing and I will type them."),
]

NOTE = ("Nothing here could be filled, because the tracker's rule is that the "
        "department sets up every job line and Finance never types a job number. "
        "Once the rows exist, send the file back and I will type the invoice "
        "number, the date and the amount on each one.")


def main():
    if len(sys.argv) < 2:
        print(__doc__.strip())
        return 1
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Still to do'
    ws.sheet_view.showGridLines = False
    for c, w in ((1, 34), (2, 13), (3, 12), (4, 26), (5, 14), (6, 12), (7, 86)):
        ws.column_dimensions[L(c)].width = w
    t = ws.cell(1, 1, 'Invoices still missing from the FY27 Revenue Tracker')
    t.font = Font(name='Arial', bold=True, size=12, color=INK)
    ws.merge_cells('A1:G1')
    n = ws.cell(2, 1, NOTE)
    n.font = Font(name='Arial', italic=True, size=9, color=INK)
    n.fill = PatternFill('solid', fgColor=AMBER)
    n.alignment = Alignment(wrap_text=True, vertical='top')
    ws.merge_cells('A2:G2')
    ws.row_dimensions[2].height = 30
    for c, h in enumerate(('Why', 'Invoice No', 'Date', 'Client', 'Net Ex GST',
                           'Job in Xero', 'What is needed'), 1):
        cell = ws.cell(4, c, h)
        cell.font = Font(name='Arial', bold=True, size=9, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor=INK)
    import datetime
    last = None
    for i, (grp, inv, date, client, net, job, need) in enumerate(LEFT):
        r = 5 + i
        ws.cell(r, 1, grp if grp != last else '')
        last = grp
        ws.cell(r, 2, inv)
        d = ws.cell(r, 3, datetime.datetime.strptime(date, '%Y-%m-%d'))
        d.number_format = 'dd/mm/yyyy'
        ws.cell(r, 4, client)
        ws.cell(r, 5, net).number_format = MONEY
        ws.cell(r, 6, job)
        w = ws.cell(r, 7, need)
        w.alignment = Alignment(wrap_text=True, vertical='top')
        for c in range(1, 8):
            ws.cell(r, c).font = Font(name='Arial', bold=(c == 1), size=9, color=INK)
        ws.row_dimensions[r].height = 28 if len(need) > 86 else 15
    r = 5 + len(LEFT)
    tot = ws.cell(r, 4, 'Total')
    tot.font = Font(name='Arial', bold=True, size=10, color=INK)
    cell = ws.cell(r, 5, '=SUM(E5:E%d)' % (r - 1))
    cell.number_format = MONEY
    cell.font = Font(name='Arial', bold=True, size=10, color=INK)
    cell.fill = PatternFill('solid', fgColor=PANEL)
    ws.freeze_panes = 'A5'
    wb.save(sys.argv[1])
    print('  %d left to do, written to %s' % (len(LEFT), sys.argv[1]))
    return 0


if __name__ == '__main__':
    sys.exit(main())
