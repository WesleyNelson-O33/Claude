"""List the invoices in Xero that are not on the FY27 Revenue Tracker.

Read from Xero through the live connection on 13 November 2026 and compared
against the tracker file dated 29 September 2026, on the invoice number. Every
month of the financial year was checked, July to date, not just the month in
front of me.

The invoice block is pasted. It does not recalculate. Every total is a formula.

    scripts/missing_invoices.py <out.xlsx>
"""
import sys

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

INK, PANEL, AMBER, PINK = '242C39', 'F4F4F4', 'FFF2CC', 'FFC7CE'
MONEY = '#,##0.00'

# number, date, client, net ex GST, status, reference
MISSING = [
 ("INV-10584","2026-08-25","MUFG Corporate Markets IR Pty Ltd",6948.50,"PAID","Orient Capital - Suncorp Results 12.08.2026 - Webcast, Cameras & vMix"),
 ("INV -10585","2026-08-25","MUFG Corporate Markets IR Pty Ltd",21703.75,"PAID","Orient Capital - Suncorp Results 12.08.2026"),
 ("INV-10652","2026-09-30","Aware Super",11000.00,"PAID","Aware Super Annual National Maintenance Proposal"),
 ("INV-10669","2026-09-30","Global Philanthropic Pty Ltd",1249.00,"AUTHORISED","Talking Philanthropy Event B-Roll"),
 ("INV-10685","2026-10-30","Commonwealth Bank of Australia (CBA Brisbane)",7434.50,"PAID","85720001BRISCSJUN27"),
 ("INV-10686","2026-10-30","Commonwealth Bank of Australia (CBA Place South)",22610.10,"PAID","85720001CBPSCSJUN27"),
 ("INV-10689","2026-10-30","PwC Services",7150.00,"AUTHORISED","PO2601702371"),
 ("INV-10691","2026-10-30","Commonwealth Bank of Australia (CBA Square)",10720.00,"PAID","85720001CBSCSJUN27"),
 ("INV-10692","2026-10-20","Dimensional Fund Advisors",380.00,"AUTHORISED","Dimensional - AV Audio Support - 15.10.26"),
 ("INV-10693","2026-10-20","Garvan Research Foundation",940.00,"PAID","Garvan - Suttons Cocktail Event 15.10.2026 (PO-58892)"),
 ("INV-10697","2026-10-30","Commonwealth Bank of Australia (CBA Melbourne)",11960.00,"PAID","85720001MEL435CSJUN27"),
 ("INV-10701","2026-10-30","PwC Services",11907.00,"AUTHORISED","PO2601702104"),
 ("INV-10703","2026-10-29","ASX Operations Pty Ltd",1730.00,"AUTHORISED","ASX - Listing Event Lighting - 14.10.26 (PO0000080953)"),
 ("INV-10704","2026-10-29","ASX Operations Pty Ltd",3674.00,"AUTHORISED","ASX - CEO Connect Series 27.10.2026"),
 ("INV-10705","2026-10-29","ASX Operations Pty Ltd",64222.51,"AUTHORISED","ASX - AGM 22.10.2026"),
 ("INV-10706","2026-10-29","APRIL 5 PTY LTD",29502.50,"AUTHORISED","April 5 - Microsoft SME&C Event - 14.10.2026 (PO016066)"),
 ("INV-10707","2026-10-29","Commonwealth Bank of Australia (CBA Place South)",3180.00,"PAID","CBA - Foldback Monitor Hire 22.10.2026"),
 ("INV-10708","2026-10-29","Elkiem",380.00,"PAID","Elkiem - Studio Support 23.10.26"),
 ("INV-10709","2026-10-29","MUFG Corporate Markets IR Pty Ltd",8299.75,"AUTHORISED","MUFG - IMB Limited AGM 27.10.2026"),
 ("INV-10710","2026-10-29","The Lottery Corporation Limited",35625.00,"AUTHORISED","TLC - Annual General Meeting 14.10.2026 (Variation)"),
 ("INV-10711","2026-10-30","Commonwealth Bank of Australia (CBA Melbourne)",680.00,"PAID","85720001MEL435CSJUN27 AFTERHOURS"),
 ("INV-10712","2026-10-30","Commonwealth Bank of Australia (CBA Place South)",1950.00,"PAID","85720001CBPSCSJUN27 AFTERHOURS"),
 ("INV-10713","2026-10-30","Commonwealth Bank of Australia (CBA Melbourne)",680.00,"PAID","85720001MEL435CSJUN27 DEDICATED TECH"),
 ("INV-10714","2026-10-30","Commonwealth Bank of Australia (CBA Place South)",425.00,"PAID","85720001CBPSCSJUN27 DEDICATED TECH"),
 ("INV-10715","2026-10-30","Commonwealth Bank of Australia (CBA Melbourne)",340.00,"PAID","85720001MEL435CSJUN27 DEDICATED TECH"),
 ("INV-10716","2026-10-30","Commonwealth Bank of Australia (CBA Melbourne)",680.00,"PAID","85720001MEL435CSJUN27 DEDICATED TECH"),
 ("INV-10717","2026-10-30","Commonwealth Bank of Australia (CBA Place South)",340.00,"PAID","85720001CBPSCSJUN27 DEDICATED TECH"),
 ("INV-10718","2026-10-30","Commonwealth Bank of Australia (CBA Place South)",425.00,"PAID","85720001CBPSCSJUN27 DEDICATED TECH"),
 ("INV-10719","2026-10-30","Commonwealth Bank of Australia (CBA Place South)",340.00,"PAID","85720001CBPSCSJUN27 DEDICATED TECH"),
 ("INV-10724","2026-10-31","Commonwealth Bank of Australia (CBA Place South)",377777.28,"AUTHORISED","CBA - Annual General Meeting 14.10.2026"),
 ("INV-10725","2026-10-30","Cochlear Limited",63911.00,"AUTHORISED","Cochlear - Annual General Meeting 26.10.2026"),
 ("INV-10726","2026-10-30","APT Management Services Pty Limited",785.00,"AUTHORISED","APA - AV Train the Trainer Session - 21.10.26"),
 ("INV-10727","2026-10-30","Commonwealth Bank of Australia (CBA Place South)",9927.50,"AUTHORISED","CBA - Live Event Support October 2026 (PO3323120)"),
 ("INV-10728","2026-10-30","MUFG Corporate Markets IR Pty Ltd",8070.50,"AUTHORISED","MUFG - Suncorp Investor Day 29.10.2026 - Webcast"),
 ("INV-10729","2026-10-30","MUFG Corporate Markets IR Pty Ltd",30317.75,"AUTHORISED","MUFG - Suncorp Investor Day 29.10.2026 - AV"),
 ("INV-10730","2026-10-30","MUFG Corporate Markets IR Pty Ltd",2000.00,"AUTHORISED","MUFG - EDV Comms 30.10.2026"),
 ("INV-10731","2026-10-30","MUFG Corporate Markets IR Pty Ltd",18796.30,"AUTHORISED","MUFG - Endeavour Annual General Meeting 30.10.2026"),
 ("INV-10732","2026-10-30","MUFG Corporate Markets IR Pty Ltd",7649.55,"AUTHORISED","MUFG - Endeavour AGM 30.10.2026 - Webcast & Cameras"),
 ("INV-10733","2026-10-30","QBE Insurance (Australia) Limited",8632.00,"AUTHORISED","QBE - Keynote Presentation 27.10.2026 (Melb)"),
 ("INV-10734","2026-10-30","The George Institute for Global Health",3360.00,"AUTHORISED","Sydney Cardiovascular Symposium 2026 29-30.10.2026"),
 ("INV-10739","2026-10-30","Bank of Western Australia Ltd (Bankwest)",1087.00,"PAID","Bankwest - Stage Hire - 29.10.26"),
 ("INV-10740","2026-10-30","Ventia",16716.80,"AUTHORISED","Ventia - CEO Town Hall (MEL) 28.10.2026 (PO4701402315)"),
 ("INV-10741","2026-10-31","Commonwealth Bank of Australia (CBA Place South)",4775.00,"AUTHORISED","CBA - Global Markets Conference 19-21.10.2026 (PO3321201)"),
]

# month, invoices in Xero, invoices on the tracker
COUNTS = [("July 2026", 43, 43), ("August 2026", 58, 56),
          ("September 2026", 52, 50), ("October 2026", 61, 22),
          ("November 2026 to date", 0, 0)]

NOTES = [
 ("Where this came from", ""),
 ("Xero", "Every sales invoice dated 1 July 2026 to 13 November 2026, read through "
          "the live Xero connection on 13 November 2026. 214 invoices."),
 ("Tracker", "FY27 Revenue Tracker 29092026.xlsx. Invoice numbers taken from the "
             "Finance, Production, Onsite and Consulting tabs."),
 ("Matched on", "Invoice number only. An invoice on the tracker with the wrong "
                "amount or the wrong month would not show up here."),
 ("This block", "Pasted from Xero. It does not recalculate. The totals are formulas."),
 ("", ""),
 ("Two other things worth a look", ""),
 ("INV -10585", "The number in Xero has a space in it, INV space 10585. There is also "
                "a real INV-10585 for PwC at 11,907.00. So Xero has two invoices that "
                "read as 10585. The PwC one is on the tracker. The MUFG one at "
                "21,703.75 is not, and the stray space is the likely reason."),
 ("INV-10521", "The tracker dates this 20 August 2026. Xero dates it 15 July 2026. "
               "Same invoice, 54,051.40 to APT Management. Worth correcting on the "
               "tracker so the month lands in July."),
 ("", ""),
 ("Not checked", ""),
 ("Amounts", "Nothing here compares the tracker amount to the Xero amount. Say the "
             "word and I will run that as well."),
 ("Credit notes", "Credit notes were not pulled. INV-10654 and INV-10642 show as fully "
                  "credited in Xero and both are on the tracker."),
]


def head(ws, r, c, text, size=11, fill=INK, colour='FFFFFF', span=None):
    cell = ws.cell(r, c, text)
    cell.font = Font(name='Arial', bold=True, size=size, color=colour)
    if fill:
        cell.fill = PatternFill('solid', fgColor=fill)
    if span:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=span)
    return cell


def main():
    if len(sys.argv) < 2:
        print(__doc__.strip())
        return 1
    out = sys.argv[1]
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = 'Summary'
    ws.sheet_view.showGridLines = False
    for c, w in ((1, 26), (2, 16), (3, 16), (4, 14), (5, 18)):
        ws.column_dimensions[L(c)].width = w
    head(ws, 1, 1, 'Invoices in Xero that are not on the FY27 Revenue Tracker', 12, span=5)
    sub = ws.cell(2, 1, 'Checked across the whole financial year, July to date, not just '
                        'the current month. Read from Xero on 13 November 2026.')
    sub.font = Font(name='Arial', italic=True, size=9, color=INK)
    ws.merge_cells('A2:E2')
    for c, h in enumerate(('Month', 'Invoices in Xero', 'On the tracker', 'Missing',
                           'Missing, net ex GST'), 1):
        head(ws, 4, c, h, 9, PANEL, INK)
    for i, (month, nx, nt) in enumerate(COUNTS):
        r = 5 + i
        ws.cell(r, 1, month).font = Font(name='Arial', size=10, color=INK)
        ws.cell(r, 2, nx)
        ws.cell(r, 3, nt)
        ws.cell(r, 4, '=B%d-C%d' % (r, r))
        ws.cell(r, 5, '=SUMPRODUCT((TEXT(\'Missing invoices\'!$B$5:$B$%d,"mmmm yyyy")=$A%d)'
                      '*\'Missing invoices\'!$D$5:$D$%d)'
                      % (4 + len(MISSING), r, 4 + len(MISSING)))
        ws.cell(r, 5).number_format = MONEY
    tot = 5 + len(COUNTS)
    head(ws, tot, 1, 'Total', 10, PANEL, INK)
    for c in (2, 3, 4):
        cell = ws.cell(tot, c, '=SUM(%s5:%s%d)' % (L(c), L(c), tot - 1))
        cell.font = Font(name='Arial', bold=True, size=10, color=INK)
        cell.fill = PatternFill('solid', fgColor=PANEL)
    cell = ws.cell(tot, 5, "=SUM('Missing invoices'!D5:D%d)" % (4 + len(MISSING)))
    cell.number_format = MONEY
    cell.font = Font(name='Arial', bold=True, size=10, color=INK)
    cell.fill = PatternFill('solid', fgColor=PANEL)
    n = ws.cell(tot + 2, 1, 'November has no invoices raised yet, so nothing is missing '
                            'there. July is clean. The work is October.')
    n.font = Font(name='Arial', italic=True, size=9, color=INK)
    ws.merge_cells(start_row=tot + 2, start_column=1, end_row=tot + 2, end_column=5)

    mi = wb.create_sheet('Missing invoices')
    mi.sheet_view.showGridLines = False
    for c, w in ((1, 13), (2, 12), (3, 46), (4, 15), (5, 13), (6, 62), (7, 15), (8, 15)):
        mi.column_dimensions[L(c)].width = w
    head(mi, 1, 1, 'Invoices in Xero with no matching invoice number on the tracker', 12, span=8)
    s = mi.cell(2, 1, 'Pasted from Xero on 13 November 2026. This block does not '
                      'recalculate. Job Number and Department are blank for you to fill in.')
    s.font = Font(name='Arial', italic=True, size=9, color=INK)
    s.fill = PatternFill('solid', fgColor=AMBER)
    mi.merge_cells('A2:H2')
    for c, h in enumerate(('Invoice No', 'Date', 'Client', 'Net Ex GST', 'Status',
                           'Reference', 'Job Number', 'Department'), 1):
        head(mi, 4, c, h, 9, PANEL, INK)
    import datetime
    for i, (num, date, client, net, status, ref) in enumerate(MISSING):
        r = 5 + i
        mi.cell(r, 1, num)
        d = mi.cell(r, 2, datetime.datetime.strptime(date, '%Y-%m-%d'))
        d.number_format = 'dd/mm/yyyy'
        mi.cell(r, 3, client)
        mi.cell(r, 4, net).number_format = MONEY
        mi.cell(r, 5, status)
        mi.cell(r, 6, ref)
        for c in range(1, 9):
            mi.cell(r, c).font = Font(name='Arial', size=9, color=INK)
        if ' ' in num.replace('INV-', '').replace('INV ', ''):
            pass
        if num == 'INV -10585':
            for c in range(1, 9):
                mi.cell(r, c).fill = PatternFill('solid', fgColor=PINK)
    r = 5 + len(MISSING)
    head(mi, r, 1, 'Total', 10, PANEL, INK)
    cell = mi.cell(r, 4, '=SUM(D5:D%d)' % (r - 1))
    cell.number_format = MONEY
    cell.font = Font(name='Arial', bold=True, size=10, color=INK)
    cell.fill = PatternFill('solid', fgColor=PANEL)
    mi.freeze_panes = 'A5'

    nt = wb.create_sheet('Notes')
    nt.sheet_view.showGridLines = False
    nt.column_dimensions['A'].width = 26
    nt.column_dimensions['B'].width = 104
    for i, (a, b) in enumerate(NOTES, 1):
        ca = nt.cell(i, 1, a)
        if b == '' and a:
            ca.font = Font(name='Arial', bold=True, size=10, color='FFFFFF')
            ca.fill = PatternFill('solid', fgColor=INK)
        else:
            ca.font = Font(name='Arial', bold=True, size=9, color=INK)
        cb = nt.cell(i, 2, b)
        cb.font = Font(name='Arial', size=9, color=INK)
        cb.alignment = Alignment(wrap_text=True, vertical='top')
        nt.row_dimensions[i].height = 42 if len(b) > 100 else (28 if len(b) > 60 else 15)

    wb.save(out)
    print('  %d missing invoices written to %s' % (len(MISSING), out))
    return 0


if __name__ == '__main__':
    sys.exit(main())
