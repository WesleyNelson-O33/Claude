"""Build the CTS weekly cash flow forecast workbook (28 Sep - 1 Nov 2026).

Data sources: Xero (cash position, aged receivables, aged payables, repeating
bills/invoices, cash-basis P&L) pulled 28 Sep 2026; FY27 GL in the Controller
Pack; the payroll manuals/transcripts in this repo.
"""
from datetime import date, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = "/home/user/Claude/cashflow/CTS Cash Flow Forecast - Oct 2026.xlsx"

F = "Arial"
BLUE = Font(name=F, size=10, color="0000FF")
BLACK = Font(name=F, size=10)
GREEN = Font(name=F, size=10, color="008000")
BOLD = Font(name=F, size=10, bold=True)
TITLE = Font(name=F, size=14, bold=True)
H2 = Font(name=F, size=11, bold=True)
GREY = Font(name=F, size=9, italic=True, color="666666")
YELLOW = PatternFill("solid", fgColor="FFFF00")
HDR = PatternFill("solid", fgColor="D9E1F2")
SUB = PatternFill("solid", fgColor="F2F2F2")
TOTAL = PatternFill("solid", fgColor="FCE4D6")
thin = Side(style="thin", color="999999")
TOP = Border(top=thin)
MONEY = '$#,##0;($#,##0);-'
MONEY2 = '$#,##0.00;($#,##0.00);-'
PCT = '0.0%'
DATE = 'ddd d mmm yy'
DATE_S = 'd mmm yy'

wb = openpyxl.Workbook()

def style_sheet(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.font is None or c.font.name != F:
                c.font = Font(name=F, size=10, bold=c.font.bold, italic=c.font.italic, color=c.font.color)

def put(ws, ref, val, font=BLACK, fmt=None, fill=None, bold=False, wrap=False):
    c = ws[ref]
    c.value = val
    f = font
    if bold:
        f = Font(name=F, size=f.size, bold=True, color=f.color, italic=f.italic)
    c.font = f
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if wrap: c.alignment = Alignment(wrap_text=True, vertical="top")
    return c

# ---------------------------------------------------------------- Assumptions
A = wb.active
A.title = "Assumptions"
A.column_dimensions["A"].width = 52
A.column_dimensions["B"].width = 16
A.column_dimensions["C"].width = 95
put(A, "A1", "CTS Cash Flow Forecast - Assumptions & Inputs", TITLE)
put(A, "A2", "Blue = input you can change. Black = formula. Yellow = key assumptions to confirm. All $ are AUD, GST inclusive (cash actually moving).", GREY)

rows = [
    ("TIMING & OPENING POSITION", None, None, None),
    ("Forecast start (Monday of this week)", date(2026, 9, 28), DATE, "Today is Mon 28 Sep 2026. Only 3 days of September remain, so the forecast runs this week plus the four full October weeks (to Sun 1 Nov)."),
    ("Opening cash - all bank accounts", 909159.55, MONEY2, "Xero cash position report, 28 Sep 2026 (cheque, savings, PL/DL credit cards etc. combined). Xero MCP does not give the split per account - payroll is paid from the cheque account with top-ups from savings, so check the cheque balance separately."),
    ("Number of weeks", 5, None, "Wk1 28 Sep-4 Oct, Wk2 5-11 Oct, Wk3 12-18 Oct, Wk4 19-25 Oct, Wk5 26 Oct-1 Nov."),
    ("PAYROLL (fortnightly, paid Tuesday of processing week)", None, None, None),
    ("Gross wages per fortnight", 111000, MONEY, "Average of the last three known pay runs: 11 Aug $108,325 and 25 Aug $115,340 (Controller Pack GL journals #69396, #69669) and the two September runs $217,062 combined (Xero cash P&L 1-28 Sep). See Payroll tab."),
    ("Net pay as % of gross", 0.77, PCT, "25 Aug run: gross ~$115k, ABA file ~$88-89k (Payroll Part 6 transcript, 03:16-04:12). 88.5/115.3 = 76.8%. The rest is PAYG withholding and salary-sacrifice deductions."),
    ("PAYG withheld as % of gross", "=1-B10", PCT, "Formula: 1 - net %. Paid to the ATO on the IAS/BAS, not on payday."),
    ("Superannuation as % of gross", 0.1275, PCT, "25 Aug run: super $14,678 / gross $115,340 = 12.7% (GL #69663). Sep: $27,816 / $217,062 = 12.8%. Above 12% SG because of salary-sacrifice super."),
    ("Pay date - run 1 (fortnight ending Fri 2 Oct)", date(2026, 10, 6), DATE, "Pay date is always the Tuesday of the payroll processing week (Payroll Part 2 transcript; Checklist step 13). Super is direct-debited by Employment Hero at the same time (Checklist step 69) - Payday Super applies from 1 Jul 2026."),
    ("Pay date - run 2 (fortnight ending Fri 16 Oct)", date(2026, 10, 20), DATE, "Next run after that is Tue 3 Nov (fortnight ending 30 Oct) - outside this forecast."),
    ("Payroll tax NSW - September return (due 7 Oct)", 6295, MONEY, "August actual accrual $6,120.40 + $174.67 (GL journals #69671, #70351). NSW monthly return and payment due the 7th of the following month. Confirm against the Revenue NSW return when lodged."),
    ("Payroll tax other states (ACT/QLD/SA/VIC/WA) - Sep (due 7 Oct)", 2700, MONEY, "August actual accrual $2,700 (GL #69670; July was $3,000). All these states are also due on the 7th of the following month."),
    ("Payroll tax payment date", date(2026, 10, 7), DATE, "7 Oct 2026 is a Wednesday."),
    ("BAS / PAYG WITHHOLDING", None, None, None),
    ("Q1 BAS paid inside this forecast? (1 = yes, 0 = no)", 0, None, "Sep-quarter BAS is due 28 Oct if self-lodged, or 25 Nov if lodged through the tax agent. The Jun-quarter BAS was paid around 25 Aug (Part 6 transcript, 04:40), which matches the tax-agent date, so the default is 0 (paid in November). Set to 1 to see the October hit."),
    ("Sep PAYG withholding (goes on the Q1 BAS)", "=B11*217062", MONEY, "Formula: PAYG % x September gross wages $217,062 (Xero cash P&L). July and August withholding were paid on the monthly IAS (21 Aug, 21 Sep)."),
    ("Net GST payable for Sep quarter - ESTIMATE", 50000, MONEY, "Rough: GST collected on ~$1.5m of quarterly receipts less GST on ~$1.0m of purchases. Replace with the figure from the Xero Activity Statement before relying on it."),
    ("PAYG income-tax instalment for Sep quarter", 0, MONEY, "Unknown - not visible through the Xero connector. Enter the ATO instalment amount if one applies."),
    ("BAS payment date if paid in October", date(2026, 10, 28), DATE, "Self-lodgement due date."),
    ("OTHER RECURRING OUTFLOWS (ESTIMATES)", None, None, None),
    ("Credit card settlements per month (PL + DL cards)", 25000, MONEY, "September transfers from cheque to cards: $15,000 (8 Sep) + $4,224.75 (17 Sep) to PL card, $5,718.88 (18 Sep) to DL card = $24,944 (Xero unreconciled bank transactions). Covers every bill marked 'paid via PL/DL CC' so those are excluded from the Payables tab totals."),
    ("Other direct costs not yet billed - per month", 100000, MONEY, "Sub-contractors, equipment/service hires, project equipment and licences that are invoiced and paid within the month. FY27 GL non-wage cost of sales: Jul $132k, Aug $147k (accrual). About 70% of that is cash within the month; the balance sits in creditors and is already listed. Budget flags October as a heavy Production month ($587k income budget) so this could run higher."),
    ("Other overheads not itemised - per month", 15000, MONEY, "Recruitment, insurance, accounting/legal, amenities, bank fees, training - items not on the Payables tab. FY27 GL non-wage overheads ran $89k (Jul) and $102k (Aug) before rent, IT and the itemised bills."),
    ("October-dated invoices collected within October", 20000, MONEY, "September evidence: invoices dated in Sep and paid in Sep (Ubank $10,867, RSNSW $3,587, CSIRO $1,971, Garvan, Elkiem, Bankwest event, St Vincent's) = ~$20k. Mostly card/Stripe payers."),
    ("Interest received on savings", 600, MONEY, "September $648.92 (Xero cash P&L). Falls on the last day of the month (week 5)."),
]
r = 4
for label, val, fmt, note in rows:
    if val is None and fmt is None and note is None:
        put(A, f"A{r}", label, H2, fill=HDR); A[f"B{r}"].fill = HDR; A[f"C{r}"].fill = HDR
    else:
        put(A, f"A{r}", label)
        is_formula = isinstance(val, str) and val.startswith("=")
        put(A, f"B{r}", val, BLACK if is_formula else BLUE, fmt)
        put(A, f"C{r}", note, GREY, wrap=True)
    r += 1
# highlight key assumptions
for ref in ("B9", "B11", "B14", "B15", "B19", "B21", "B26", "B27"):
    A[ref].fill = YELLOW
# named positions used elsewhere
AS = {
    "start": "Assumptions!$B$5", "opening": "Assumptions!$B$6", "gross": "Assumptions!$B$9",
    "netpct": "Assumptions!$B$10", "paygpct": "Assumptions!$B$11", "superpct": "Assumptions!$B$12",
    "pay1": "Assumptions!$B$13", "pay2": "Assumptions!$B$14", "ptnsw": "Assumptions!$B$15",
    "ptoth": "Assumptions!$B$16", "ptdate": "Assumptions!$B$17", "bastoggle": "Assumptions!$B$19",
    "baspayg": "Assumptions!$B$20", "basgst": "Assumptions!$B$21", "basinst": "Assumptions!$B$22",
    "basdate": "Assumptions!$B$23", "cc": "Assumptions!$B$25", "directcost": "Assumptions!$B$26",
    "overhead": "Assumptions!$B$27", "quickpay": "Assumptions!$B$28", "interest": "Assumptions!$B$29",
}
# sanity: check labels line up with the map
assert A["A5"].value.startswith("Forecast start"), A["A5"].value
assert A["A6"].value.startswith("Opening cash")
assert A["A9"].value.startswith("Gross wages")
assert A["A10"].value.startswith("Net pay")
assert A["A11"].value.startswith("PAYG withheld")
assert A["A12"].value.startswith("Superannuation")
assert A["A13"].value.startswith("Pay date - run 1")
assert A["A14"].value.startswith("Pay date - run 2")
assert A["A15"].value.startswith("Payroll tax NSW")
assert A["A16"].value.startswith("Payroll tax other")
assert A["A17"].value.startswith("Payroll tax payment date")
assert A["A19"].value.startswith("Q1 BAS paid")
assert A["A20"].value.startswith("Sep PAYG")
assert A["A21"].value.startswith("Net GST")
assert A["A22"].value.startswith("PAYG income")
assert A["A23"].value.startswith("BAS payment date")
assert A["A25"].value.startswith("Credit card")
assert A["A26"].value.startswith("Other direct")
assert A["A27"].value.startswith("Other overheads")
assert A["A28"].value.startswith("October-dated")
assert A["A29"].value.startswith("Interest")
# fix the yellow highlights to the right cells now that rows are known
for ref in ("B9", "B11", "B14", "B15", "B19", "B21", "B26", "B27"):
    A[ref].fill = PatternFill(fill_type=None)
for ref in ("B9", "B10", "B15", "B16", "B19", "B21", "B26", "B27"):
    A[ref].fill = YELLOW

# Week table
put(A, "A31", "WEEK TABLE (calculated)", H2, fill=HDR); A["B31"].fill = HDR; A["C31"].fill = HDR
put(A, "A32", "Week", BOLD); put(A, "B32", "Start (Mon)", BOLD); put(A, "C32", "End (Sun)", BOLD)
for i in range(5):
    rr = 33 + i
    put(A, f"A{rr}", i + 1)
    put(A, f"B{rr}", f"={AS['start']}+{7*i}", BLACK, DATE)
    put(A, f"C{rr}", f"=B{rr}+6", BLACK, DATE)
WEEK_START = [f"Assumptions!$B${33+i}" for i in range(5)]
WEEK_END = [f"Assumptions!$C${33+i}" for i in range(5)]
FC_END = "Assumptions!$C$37"

# ---------------------------------------------------------------- Payroll
P = wb.create_sheet("Payroll")
for col, w in zip("ABCDEFGHI", (30, 16, 16, 14, 14, 14, 14, 14, 60)):
    P.column_dimensions[col].width = w
put(P, "A1", "Payroll schedule - October 2026", TITLE)
put(P, "A2", "Fortnightly pay cycle. Pay date = Tuesday of the processing week. Super is direct-debited in the same week. PAYG withholding is paid later via IAS/BAS, not on payday.", GREY)
hdr = ["Pay run", "Fortnight ending", "Pay date", "Gross wages", "Net pay (ABA file)", "PAYG withheld", "Super (direct debit)", "Cash out on payday (net + super)", "Notes"]
for j, h in enumerate(hdr):
    put(P, f"{get_column_letter(j+1)}4", h, BOLD, fill=HDR, wrap=True)
runs = [("Run 1 - Oct", date(2026, 10, 2), AS["pay1"], "Timesheets close Fri 2 Oct, prep starts that Friday, processed Mon 5 Oct, paid Tue 6 Oct."),
        ("Run 2 - Oct", date(2026, 10, 16), AS["pay2"], "Last pay run of the month: accrue bonus leave (Checklist step 48) - no cash effect.")]
for i, (name, fe, payref, note) in enumerate(runs):
    rr = 5 + i
    put(P, f"A{rr}", name)
    put(P, f"B{rr}", fe, BLUE, DATE)
    put(P, f"C{rr}", f"={payref}", GREEN, DATE)
    put(P, f"D{rr}", f"={AS['gross']}", GREEN, MONEY)
    put(P, f"E{rr}", f"=D{rr}*{AS['netpct']}", BLACK, MONEY)
    put(P, f"F{rr}", f"=D{rr}*{AS['paygpct']}", BLACK, MONEY)
    put(P, f"G{rr}", f"=D{rr}*{AS['superpct']}", BLACK, MONEY)
    put(P, f"H{rr}", f"=E{rr}+G{rr}", BLACK, MONEY)
    put(P, f"I{rr}", note, GREY, wrap=True)
put(P, "A7", "Total October pay runs", BOLD, fill=TOTAL)
for col in "DEFGH":
    put(P, f"{col}7", f"=SUM({col}5:{col}6)", BOLD, MONEY, fill=TOTAL)

put(P, "A9", "Payroll tax for September wages (paid October)", H2)
put(P, "A10", "NSW - due 7 Oct"); put(P, "B10", f"={AS['ptnsw']}", GREEN, MONEY); put(P, "C10", f"={AS['ptdate']}", GREEN, DATE)
put(P, "A11", "Other states - due 7 Oct"); put(P, "B11", f"={AS['ptoth']}", GREEN, MONEY); put(P, "C11", f"={AS['ptdate']}", GREEN, DATE)
put(P, "A12", "Total payroll tax", BOLD); put(P, "B12", "=B10+B11", BOLD, MONEY)
put(P, "A13", "Payroll tax for October wages is due 7 Nov - outside this forecast.", GREY)

put(P, "A15", "Reference - actual pay runs this financial year", H2)
for j, h in enumerate(["Pay run", "Pay date", "Gross wages", "Super", "Super % of gross", "Source"]):
    put(P, f"{get_column_letter(j+1)}16", h, BOLD, fill=HDR)
actuals = [
    ("FE 10 Jul", date(2026, 7, 14), 104335.03, 13289.41, "Controller Pack GL_Paste journals #69527 / #69434"),
    ("FE 24 Jul", date(2026, 7, 28), 101030.18, 13068.87, "GL journals #68743 / #68744"),
    ("FE 7 Aug", date(2026, 8, 11), 108324.58, 13783.25, "GL journals #69396 / #69397"),
    ("FE 21 Aug", date(2026, 8, 25), 115339.80, 14677.93, "GL journals #69669 / #69663 (the run shown in the training videos)"),
    ("FE 4 Sep + FE 18 Sep (two runs)", date(2026, 9, 22), 217061.76, 27815.63, "Xero cash-basis P&L 1-28 Sep 2026: all Direct/Indirect Salaries & Wages accounts; Direct + Indirect Superannuation"),
]
for i, (n, d, g, s, src) in enumerate(actuals):
    rr = 17 + i
    put(P, f"A{rr}", n); put(P, f"B{rr}", d, BLUE, DATE); put(P, f"C{rr}", g, BLUE, MONEY); put(P, f"D{rr}", s, BLUE, MONEY)
    put(P, f"E{rr}", f"=IF(C{rr}=0,0,D{rr}/C{rr})", BLACK, PCT); put(P, f"F{rr}", src, GREY)
put(P, "A22", "Average gross per fortnight (last 4 fortnights)", BOLD)
put(P, "B22", "=(C19+C20+C21)/4", BOLD, MONEY)
put(P, "C22", "Aug runs plus the two Sep runs, divided by four fortnights", GREY)
put(P, "A24", "Reference - payroll tax accruals this financial year", H2)
for j, h in enumerate(["Month", "NSW", "Other states", "Total", "Source"]):
    put(P, f"{get_column_letter(j+1)}25", h, BOLD, fill=HDR)
put(P, "A26", "July 2026"); put(P, "B26", 5170.80, BLUE, MONEY); put(P, "C26", 3000, BLUE, MONEY); put(P, "D26", "=B26+C26", BLACK, MONEY); put(P, "E26", "GL journals #68926 (NSW) and #68927 (other states)", GREY)
put(P, "A27", "August 2026"); put(P, "B27", 6295.07, BLUE, MONEY); put(P, "C27", 2700, BLUE, MONEY); put(P, "D27", "=B27+C27", BLACK, MONEY); put(P, "E27", "GL journals #69671 + #70351 (NSW) and #69670 (other states)", GREY)
put(P, "A28", "Note: FY27 budget has payroll tax at ~$7.4k/month; actuals are running ~$9k/month, so actuals are used.", GREY)

# ---------------------------------------------------------------- Receivables
R = wb.create_sheet("Receivables")
widths = (34, 12, 44, 12, 12, 14, 13, 16, 11, 8, 60)
for j, w in enumerate(widths):
    R.column_dimensions[get_column_letter(j+1)].width = w
put(R, "A1", "Receivables - every open invoice and when the cash is expected", TITLE)
put(R, "A2", "Source: Xero aged receivables 28 Sep 2026 ($580,292.39, 38 invoices) plus the September contract invoices still to be raised (Xero repeating invoice templates). Expected date (blue) is a judgement based on each customer's recent paying pattern - change it and the forecast moves.", GREY)
hdr = ["Customer", "Invoice", "Reference", "Invoice date", "Due date", "Amount (incl GST)", "Status", "Expected receipt date", "Days after due", "Week", "Basis for expected date"]
for j, h in enumerate(hdr):
    put(R, f"{get_column_letter(j+1)}4", h, BOLD, fill=HDR, wrap=True)
D = date
ar = [
 ("APT Management Services (APA)", "INV-10622", "PO-00030884", D(2026,9,10), D(2026,10,10), 4352.45, "Current", D(2026,10,9), "APA paid its last invoice the day after issue"),
 ("APT Management Services (APA)", "INV-10628", "APA RUOK Event 10.09.2026", D(2026,9,16), D(2026,10,16), 627.00, "Current", D(2026,10,16), "On due date"),
 ("ASX Operations", "INV-10599", "ASX Employee Town Hall 13.08.2026", D(2026,8,27), D(2026,9,26), 1661.00, "Overdue", D(2026,10,2), "ASX paid its Aug invoices ~4 days late"),
 ("ASX Operations", "INV-10624", "ASX CEO Connect Series 01.09.2026", D(2026,9,15), D(2026,10,15), 4591.40, "Current", D(2026,10,16), "Due date + a day"),
 ("ASX Operations", "INV-10623", "ASX Live 01.09.2026", D(2026,9,15), D(2026,10,15), 8544.80, "Current", D(2026,10,16), "Due date + a day"),
 ("Australian Payments Network", "INV-10574", "Logitech room licence renewal", D(2026,8,20), D(2026,9,19), 1092.30, "Overdue", D(2026,10,2), "Small, overdue - chase this week"),
 ("Aware Super", "INV-10583", "Annual National Maintenance (Aug)", D(2026,8,31), D(2026,9,30), 12100.00, "Current", D(2026,9,30), "Aware pays on or near due date"),
 ("Aware Super", "INV-10596", "Seminar Room AV Upgrade", D(2026,9,9), D(2026,10,9), 225887.40, "Current", D(2026,10,9), "KEY RECEIPT. Complete AV bill of $178,119 is on hold until this lands. If Aware slips, hold Complete AV."),
 ("Aware Super", "INV-10635", "Appspace licensing subscription (12 months)", D(2026,9,18), D(2026,10,18), 23317.11, "Current", D(2026,10,16), "Due Sun 18 Oct - expect Fri 16 Oct. Funds the Appspace bill ($20,985) due 24 Oct."),
 ("Aware Super", "INV-10646", "Perth office BYOD USB-C replacement", D(2026,9,24), D(2026,10,24), 3252.70, "Current", D(2026,10,23), "Due Sat 24 Oct - expect Fri 23 Oct"),
 ("Aware Super", "INV-10619", "Sydney 24.16 UC Engine remediation variation", D(2026,9,25), D(2026,10,25), 738.10, "Current", D(2026,10,23), "Due Sun 25 Oct - expect Fri 23 Oct"),
 ("Downer EDI", "INV-10587", "Full Year Results Briefing 20.08.2026", D(2026,8,25), D(2026,9,24), 12808.69, "Overdue", D(2026,10,2), "4 days overdue - assume paid this week"),
 ("Downer EDI", "INV-10611", "Full Year Employee Briefing 25.08.2026", D(2026,8,31), D(2026,9,30), 26855.13, "Current", D(2026,10,7), "Due 30 Sep; allow a week"),
 ("Generation-e Productivity Solutions", "INV-10532", "PAU26-407 (Jul)", D(2026,7,31), D(2026,8,30), 13627.68, "Overdue", D(2026,10,7), "A month overdue - chase. Assume paid week 2"),
 ("Generation-e Productivity Solutions", "INV-10573", "PAU26-407 (Aug)", D(2026,8,31), D(2026,9,30), 13627.68, "Current", D(2026,10,21), "Slow payer - assume 3 weeks after due"),
 ("ICC Sydney", "INV-10630", "Speaker presentation technicians - ICAS 2026", D(2026,9,16), D(2026,10,16), 1040.88, "Current", D(2026,10,16), "On due date"),
 ("ICC Sydney", "INV-10633", "Speaker presentation technicians - ISE", D(2026,9,16), D(2026,10,16), 5131.51, "Current", D(2026,10,16), "On due date"),
 ("Jaazaniah Salanoa (staff)", "INV-10621", "Goget hire recovery", D(2026,8,31), D(2026,9,30), 66.64, "Current", D(2026,10,6), "Staff recovery - likely via payroll"),
 ("Meet Magic", "INV-10464", "Human AI Event highlight videos 29.05.2026", D(2026,6,19), D(2026,6,26), 5075.07, "Overdue 3 months", D(2026,11,30), "DOUBTFUL - 3 months overdue. Excluded from October. Needs escalation."),
 ("Microsoft", "INV-10554", "PO 10160645", D(2026,7,31), D(2026,8,30), 13845.70, "Overdue", D(2026,10,7), "A month overdue - usually a PO/portal issue. Assume week 2"),
 ("ON24", "INV-10552", "PO #25962 (Jul)", D(2026,7,31), D(2026,8,30), 2832.50, "Overdue", D(2026,10,7), "A month overdue - chase"),
 ("ON24", "INV-10589", "PO25962 (Aug)", D(2026,8,26), D(2026,9,25), 3005.75, "Overdue", D(2026,10,7), "Assume paid with the July one"),
 ("Property NSW", "INV-10396", "Penrith BDA Condeco desk booking (PO20285637)", D(2026,5,6), D(2026,6,5), 19733.45, "Overdue 3 months", D(2026,11,30), "DOUBTFUL - almost 4 months overdue. Excluded from October. Needs escalation."),
 ("PwC Services", "INV-10535", "PO2601702104 (Jul)", D(2026,7,31), D(2026,8,30), 13097.70, "Overdue", D(2026,10,7), "A month overdue. PwC paid INV-10577 on its due date, so these two look like PO problems - chase"),
 ("PwC Services", "INV-10531", "PO2601702371 (Jul)", D(2026,7,31), D(2026,8,30), 7865.00, "Overdue", D(2026,10,7), "As above"),
 ("PwC Services", "INV-10598", "PO2601703509", D(2026,8,31), D(2026,9,30), 9350.00, "Current", D(2026,10,14), "Allow two weeks after due"),
 ("PwC Services", "INV-10572", "PO2601702371 (Aug)", D(2026,8,31), D(2026,9,30), 7865.00, "Current", D(2026,10,14), "Allow two weeks after due"),
 ("PwC Services", "INV-10585", "PO2601702104 (Aug)", D(2026,8,31), D(2026,9,30), 13097.70, "Current", D(2026,10,14), "Allow two weeks after due"),
 ("Ricoh Australia", "INV-10618", "CONTRACT (Aug)", D(2026,8,31), D(2026,9,30), 13046.62, "Current", D(2026,10,7), "Allow a week"),
 ("Ricoh Australia", "INV-10620", "SHELL AV Support", D(2026,9,3), D(2026,10,3), 374.00, "Current", D(2026,10,7), "Allow a week"),
 ("The CEO Institute", "INV-10581", "CEO Connect 2026 11.8.26", D(2026,8,21), D(2026,9,20), 9086.39, "Overdue", D(2026,10,2), "8 days overdue - assume this week"),
 ("Vega Global Australia", "INV-10647", "Payments & Securities Forums 14-17.9.26", D(2026,9,22), D(2026,10,22), 1969.00, "Current", D(2026,10,22), "On due date"),
 ("Ventia", "INV-10557", "Ventia Awards 29-30.7.2026 (PO4701355678)", D(2026,7,31), D(2026,8,30), 58024.76, "Overdue", D(2026,10,9), "LARGEST OVERDUE - a month late. Chase now; assume week 2"),
 ("Ventia", "INV-10612", "Half Year Results Briefing 24.08.2026", D(2026,8,31), D(2026,9,30), 11919.04, "Current", D(2026,10,21), "Ventia is paying ~3 weeks late"),
 ("Ventia", "INV-10593", "Employee HYR Town Hall 24.08.2026", D(2026,8,31), D(2026,9,30), 14741.65, "Current", D(2026,10,21), "Ventia is paying ~3 weeks late"),
 ("Ventia", "INV-10632", "RUOK? Day Event 10.09.2026", D(2026,9,16), D(2026,10,16), 10015.89, "Current", D(2026,10,30), "Two weeks after due"),
 ("YSG Studio", "INV-10601", "Orange Room AV upgrade - labour", D(2026,8,27), D(2026,9,26), 4908.20, "Overdue", D(2026,10,2), "2 days overdue - assume this week"),
 ("YSG Studio", "INV-10653", "YSG Studio Phase 4", D(2026,9,25), D(2026,10,25), 1116.50, "Current", D(2026,10,30), "Allow a few days after due"),
]
assert abs(sum(x[5] for x in ar) - 580292.39) < 0.01, sum(x[5] for x in ar)
to_raise = [
 ("Deloitte Services Trust", "TBR", "CONTRACT - September technicians", D(2026,9,30), D(2026,10,30), 113987.99, "To be raised", D(2026,10,14), "Aug invoice (31 Aug) was paid 14 Sep - Deloitte pays ~2 weeks after issue. Template: repeating invoice 'CONTRACT' $113,987.99/month"),
 ("Commonwealth Bank (CBA Place South)", "TBR", "85720001CBPSCSJUN27 - September", D(2026,9,30), D(2026,10,30), 24978.80, "To be raised", D(2026,10,8), "Aug invoices (31 Aug) paid 8 Sep - CBA pays ~8 days after issue"),
 ("Commonwealth Bank (CBA Square)", "TBR", "85720001CBSCSJUN27 - September", D(2026,9,30), D(2026,10,30), 11792.00, "To be raised", D(2026,10,8), "As above"),
 ("Commonwealth Bank (CBA Melbourne)", "TBR", "85720001MEL435CSJUN27 - September", D(2026,9,30), D(2026,10,30), 13156.00, "To be raised", D(2026,10,8), "As above"),
 ("Commonwealth Bank (CBA Brisbane)", "TBR", "85720001BRISCSJUN27 - September", D(2026,9,30), D(2026,10,30), 6534.00, "To be raised", D(2026,10,8), "As above"),
 ("Bankwest", "TBR", "85720001BWPCSJUN27 - September", D(2026,9,30), D(2026,10,30), 21210.75, "To be raised", D(2026,10,5), "Aug invoice (31 Aug) paid 3 Sep - Bankwest pays within days"),
 ("Bankwest", "TBR", "85720001BWPCSJUN27 Additionals - September", D(2026,9,30), D(2026,10,30), 1490.50, "To be raised", D(2026,10,5), "As above (template amount; actual varies)"),
 ("PwC Services", "TBR", "PO2601702104 - Sydney AV support September", D(2026,9,30), D(2026,10,30), 13097.70, "To be raised", D(2026,11,14), "PwC's Jul contract invoices are still unpaid - assume November"),
 ("PwC Services", "TBR", "PO2601702371 - Brisbane AV support September", D(2026,9,30), D(2026,10,30), 7865.00, "To be raised", D(2026,11,14), "As above"),
 ("Aware Super", "TBR", "Annual National Maintenance - September", D(2026,9,30), D(2026,10,30), 12100.00, "To be raised", D(2026,10,30), "Aware pays on due date"),
]
rr = 5
first = rr
for row in ar + to_raise:
    cust, inv, ref, idate, due, amt, status, exp, basis = row
    put(R, f"A{rr}", cust); put(R, f"B{rr}", inv); put(R, f"C{rr}", ref)
    put(R, f"D{rr}", idate, BLUE, DATE_S); put(R, f"E{rr}", due, BLUE, DATE_S)
    put(R, f"F{rr}", amt, BLUE, MONEY2); put(R, f"G{rr}", status, BLUE)
    put(R, f"H{rr}", exp, BLUE, DATE_S, fill=YELLOW if amt > 50000 else None)
    put(R, f"I{rr}", f"=H{rr}-E{rr}", BLACK, '0')
    put(R, f"J{rr}", f'=IF(H{rr}<{AS["start"]},"Before",IF(H{rr}>{FC_END},"After",INT((H{rr}-{AS["start"]})/7)+1))', BLACK)
    put(R, f"K{rr}", basis, GREY, wrap=True)
    rr += 1
last = rr - 1
put(R, f"A{rr}", "TOTAL", BOLD, fill=TOTAL)
put(R, f"F{rr}", f"=SUM(F{first}:F{last})", BOLD, MONEY2, fill=TOTAL)
put(R, f"A{rr+1}", "of which open in Xero at 28 Sep (should equal $580,292.39)", GREY)
put(R, f"F{rr+1}", f'=SUMIFS(F{first}:F{last},G{first}:G{last},"<>To be raised")', BLACK, MONEY2)
put(R, f"A{rr+2}", "of which expected inside the forecast window", GREY)
put(R, f"F{rr+2}", f'=SUMIFS(F{first}:F{last},H{first}:H{last},">="&{AS["start"]},H{first}:H{last},"<="&{FC_END})', BLACK, MONEY2)
put(R, f"A{rr+3}", "of which pushed past 1 Nov (PwC Sep invoices, Meet Magic, Property NSW, anything you re-date)", GREY)
put(R, f"F{rr+3}", f"=F{rr}-F{rr+2}", BLACK, MONEY2)
R.freeze_panes = "A5"
R.auto_filter.ref = f"A4:K{last}"
AR_RANGE = dict(amt=f"Receivables!$F${first}:$F${last}", exp=f"Receivables!$H${first}:$H${last}", status=f"Receivables!$G${first}:$G${last}")

# ---------------------------------------------------------------- Payables
Y = wb.create_sheet("Payables")
widths = (34, 46, 12, 14, 11, 16, 8, 60)
for j, w in enumerate(widths):
    Y.column_dimensions[get_column_letter(j+1)].width = w
put(Y, "A1", "Payables - every open bill, plus October's recurring bills not yet in Xero", TITLE)
put(Y, "A2", "Source: Xero aged payables 28 Sep 2026 ($283,042.53, 24 bills) and Xero repeating bill templates. Bills marked 'Card' are paid on the PL/DL credit cards and are covered by the credit-card settlement line on the Cash Flow tab, so they are excluded from the supplier totals to avoid double counting.", GREY)
hdr = ["Supplier", "Bill / reference", "Due date", "Amount (incl GST)", "Method", "Planned payment date", "Week", "Notes"]
for j, h in enumerate(hdr):
    put(Y, f"{get_column_letter(j+1)}4", h, BOLD, fill=HDR, wrap=True)
ap = [
 ("Appspace, Inc.", "INV00133660 - Appspace licensing for Aware Super", D(2026,10,24), 20985.43, "EFT", D(2026,10,23), "Pay once Aware's matching invoice INV-10635 ($23,317) is received (expected 16 Oct)"),
 ("Complete AV Solutions", "INV41565 - Aware Super seminar room audio upgrade", D(2026,9,26), 178119.25, "EFT", D(2026,10,14), "ON HOLD per bill note 'hold off until payment received from Aware'. Aware INV-10596 ($225,887) expected 9 Oct - pay the week after. Largest single outflow."),
 ("Crestron ANZ", "INV950936412 - Aware SYD 28.06 equipment", D(2026,10,19), 7103.26, "EFT", D(2026,10,19), "On due date"),
 ("Employsure", "41 of 60", D(2026,9,6), 880.00, "EFT", D(2026,10,1), "Overdue - pay this week"),
 ("Energy Australia", "INV260382065288 - energy 23 Jul-22 Aug", D(2026,9,29), 377.13, "Card", D(2026,9,29), "Direct debit via PL credit card"),
 ("Energy Australia", "INV260412412804 - electricity 23 Aug-21 Sep", D(2026,10,14), 696.25, "Card", D(2026,10,14), "Direct debit via PL credit card"),
 ("Eptura Australia", "INV-85230 - Proxyclick renewal APA VMS yr 2", D(2026,10,1), 3627.04, "EFT", D(2026,10,1), "On due date"),
 ("Expedia", "ITN73550041732636 - accommodation D Hurley", D(2026,9,22), 808.95, "Card", D(2026,9,22), "Earlier Expedia bookings went on the PL card"),
 ("Fredon Technology", "INV11S-26080021 - DTTL 8PSQ Dream Big microphones", D(2026,9,23), 24708.20, "EFT", D(2026,10,7), "Overdue; bill note 'DH to confirm'. Assume released week 2"),
 ("Investa Asset Management", "INV5310034149 - Rent September 2026 (incl. pest)", D(2026,10,22), 29280.72, "EFT", D(2026,10,22), "On due date. October rent will be invoiced ~22 Oct and due ~22 Nov (outside forecast)"),
 ("Mailchimp", "MC27286583 - September", D(2026,9,18), 37.62, "Card", D(2026,9,18), "PL card"),
 ("Optus", "INV000594614144 - office internet & landline Sep", D(2026,9,23), 629.66, "Card", D(2026,9,23), "PL card"),
 ("Optus", "INV000594999543 - staff mobiles Sep", D(2026,10,12), 804.11, "BPAY", D(2026,10,12), "BPAY on due date"),
 ("Persona Health", "INV-18315 - fitness for duty assessment", D(2026,9,10), 996.60, "EFT", D(2026,10,21), "Bill note: GC advised hold payment until he confirms. Assumed released week 4"),
 ("Posh Services (Ambius)", "INV21736844 - indoor plants Sep", D(2026,10,9), 245.00, "EFT", D(2026,10,9), "On due date"),
 ("Ricoh Australia", "INV15759052 - printing Aug", D(2026,9,30), 235.60, "DD", D(2026,9,30), "Direct debit"),
 ("RJW Group Holdings", "INV-0048 - website design balance", D(2026,9,30), 6600.00, "EFT", D(2026,10,2), "Bill note: check with DL before payment"),
 ("Seek", "INV702112646 - job ad casual AV concierge", D(2026,9,18), 550.00, "Card", D(2026,9,18), "PL card"),
 ("SPACERA", "INV-0056 - Aware Super national support (Sep)", D(2026,10,1), 5728.80, "EFT", D(2026,10,1), "On due date"),
 ("Tec Art", "INV136603 - PRD asset purchase", D(2026,9,22), 657.50, "Card", D(2026,9,22), "PL card"),
 ("Tec Art", "INV136603CR - refund", D(2026,9,24), -657.50, "Card", D(2026,9,24), "Refund to PL card"),
 ("Transport for NSW", "Opal top-up T Wood", D(2026,9,18), 20.00, "Card", D(2026,9,18), "Card"),
 ("Virgin Australia", "SYMUMZ-3 - rescheduled flights D Hurley", D(2026,9,18), 501.91, "Card", D(2026,9,18), "Earlier Virgin bookings went on the PL card"),
 ("Xero", "INV-56302114 - subscription Sep", D(2026,9,24), 107.00, "Card", D(2026,9,24), "PL card"),
]
assert abs(sum(x[3] for x in ap) - 283042.53) < 0.01, sum(x[3] for x in ap)
recurring = [
 ("First Focus IT", "CORE Intelligent Managed Services - October", D(2026,10,16), 13755.77, "DD", D(2026,10,16), "Repeating bill template, direct debit"),
 ("First Focus IT", "Managed Services Azure - October", D(2026,10,16), 710.50, "DD", D(2026,10,16), "Repeating bill template, direct debit"),
 ("RJW Group Holdings", "Digital growth marketing - October", D(2026,10,21), 4675.00, "EFT", D(2026,10,21), "Repeating bill template"),
 ("Cleveland Partners (8020 Advisors)", "Monthly advisory fees", D(2026,10,7), 3850.00, "EFT", D(2026,10,7), "Repeating bill template $3,850/month (Sep cash P&L shows $5,000 business advisory). Date assumed"),
 ("SPACERA", "Aware Super national support - October", D(2026,10,31), 5728.80, "EFT", D(2026,10,31), "Repeating bill, next due 31 Oct"),
 ("LiveU", "SOLO standard data pack - October", D(2026,10,14), 1485.00, "EFT", D(2026,10,14), "Repeating bill, due 14 Oct"),
 ("Employsure", "42 of 60 - October", D(2026,10,5), 880.00, "EFT", D(2026,10,5), "Repeating bill, due 5 Oct"),
 ("The CEO Institute", "Syndicate membership (DL) - quarterly", D(2026,10,8), 2310.00, "EFT", D(2026,10,8), "Repeating bill every 3 months, due 8 Oct"),
 ("Telstra", "Nighthawk data plans - October", D(2026,10,31), 130.00, "EFT", D(2026,10,31), "Repeating bill"),
 ("Employment Hero", "Subscription - September", D(2026,10,30), 2048.20, "Card", D(2026,9,30), "PL card - in credit card settlement line"),
 ("Goget Carshare", "September usage", D(2026,10,30), 1783.07, "Card", D(2026,9,30), "PL card - in credit card settlement line"),
 ("Zoom", "Monthly subscription", D(2026,10,28), 323.97, "Card", D(2026,10,28), "PL card"),
 ("Insphire (Current RMS)", "October", D(2026,10,16), 431.00, "Card", D(2026,10,16), "PL card"),
 ("Optus", "Office internet & landline - October", D(2026,10,21), 629.66, "Card", D(2026,10,21), "PL card"),
 ("Adobe / Spotify / Avangate / Telstra recharges / misc", "Small monthly subscriptions", D(2026,10,15), 700.00, "Card", D(2026,10,15), "PL/DL cards - approx"),
 ("Investa Asset Management", "Rent October 2026", D(2026,11,22), 29280.72, "EFT", D(2026,11,22), "Invoiced ~22 Oct, due ~22 Nov - AFTER this forecast"),
 ("Optus", "Staff mobiles - October", D(2026,11,8), 804.06, "BPAY", D(2026,11,8), "Due 8 Nov - AFTER this forecast"),
]
rr = 5
ap_first = rr
for row in ap:
    sup, ref, due, amt, method, pdate, note = row
    put(Y, f"A{rr}", sup); put(Y, f"B{rr}", ref); put(Y, f"C{rr}", due, BLUE, DATE_S)
    put(Y, f"D{rr}", amt, BLUE, MONEY2); put(Y, f"E{rr}", method, BLUE)
    put(Y, f"F{rr}", pdate, BLUE, DATE_S, fill=YELLOW if amt > 50000 else None)
    put(Y, f"G{rr}", f'=IF(F{rr}<{AS["start"]},"Before",IF(F{rr}>{FC_END},"After",INT((F{rr}-{AS["start"]})/7)+1))', BLACK)
    put(Y, f"H{rr}", note, GREY, wrap=True)
    rr += 1
ap_last = rr - 1
put(Y, f"A{rr}", "TOTAL open bills in Xero (should equal $283,042.53)", BOLD, fill=TOTAL)
put(Y, f"D{rr}", f"=SUM(D{ap_first}:D{ap_last})", BOLD, MONEY2, fill=TOTAL)
put(Y, f"A{rr+1}", "of which paid by card (excluded - see credit card line)", GREY)
put(Y, f"D{rr+1}", f'=SUMIFS(D{ap_first}:D{ap_last},E{ap_first}:E{ap_last},"Card")', BLACK, MONEY2)
rr += 3
put(Y, f"A{rr}", "OCTOBER RECURRING BILLS NOT YET IN XERO (from repeating-bill templates)", H2, fill=HDR)
for col in "BCDEFGH": Y[f"{col}{rr}"].fill = HDR
rr += 1
rec_first = rr
for row in recurring:
    sup, ref, due, amt, method, pdate, note = row
    put(Y, f"A{rr}", sup); put(Y, f"B{rr}", ref); put(Y, f"C{rr}", due, BLUE, DATE_S)
    put(Y, f"D{rr}", amt, BLUE, MONEY2); put(Y, f"E{rr}", method, BLUE); put(Y, f"F{rr}", pdate, BLUE, DATE_S)
    put(Y, f"G{rr}", f'=IF(F{rr}<{AS["start"]},"Before",IF(F{rr}>{FC_END},"After",INT((F{rr}-{AS["start"]})/7)+1))', BLACK)
    put(Y, f"H{rr}", note, GREY, wrap=True)
    rr += 1
rec_last = rr - 1
put(Y, f"A{rr}", "TOTAL recurring (non-card, inside forecast window)", BOLD, fill=TOTAL)
put(Y, f"D{rr}", f'=SUMIFS(D{rec_first}:D{rec_last},E{rec_first}:E{rec_last},"<>Card",F{rec_first}:F{rec_last},">="&{AS["start"]},F{rec_first}:F{rec_last},"<="&{FC_END})', BOLD, MONEY2, fill=TOTAL)
Y.freeze_panes = "A5"
AP = dict(amt=f"Payables!$D${ap_first}:$D${ap_last}", pd=f"Payables!$F${ap_first}:$F${ap_last}", m=f"Payables!$E${ap_first}:$E${ap_last}", sup=f"Payables!$A${ap_first}:$A${ap_last}")
REC = dict(amt=f"Payables!$D${rec_first}:$D${rec_last}", pd=f"Payables!$F${rec_first}:$F${rec_last}", m=f"Payables!$E${rec_first}:$E${rec_last}")

# ---------------------------------------------------------------- Cash Flow
C = wb.create_sheet("Cash Flow", 0)
C.column_dimensions["A"].width = 52
for col in "BCDEF": C.column_dimensions[col].width = 15
C.column_dimensions["G"].width = 16
C.column_dimensions["H"].width = 70
put(C, "A1", "Corporate Technology Services Pty Ltd - Weekly Cash Flow Forecast", TITLE)
put(C, "A2", "This week (w/c 28 Sep 2026) through the end of October. Built from Xero data as at 28 Sep 2026. Inputs on the Assumptions tab; invoice/bill detail on the Receivables and Payables tabs.", GREY)
put(C, "A4", "Week", BOLD, fill=HDR)
put(C, "A5", "Week starting (Mon)", BOLD, fill=HDR)
put(C, "A6", "Week ending (Sun)", BOLD, fill=HDR)
for i in range(5):
    col = get_column_letter(2 + i)
    put(C, f"{col}4", f"Week {i+1}", BOLD, fill=HDR)
    put(C, f"{col}5", f"={WEEK_START[i]}", GREEN, DATE, fill=HDR)
    put(C, f"{col}6", f"={WEEK_END[i]}", GREEN, DATE, fill=HDR)
put(C, "G4", "Total 5 weeks", BOLD, fill=HDR); C["G5"].fill = HDR; C["G6"].fill = HDR
put(C, "H4", "Basis / source", BOLD, fill=HDR); C["H5"].fill = HDR; C["H6"].fill = HDR
for col in "BCDEFG":
    C[f"{col}4"].alignment = Alignment(horizontal="center")

def week_cols():
    return [get_column_letter(2 + i) for i in range(5)]

def sumifs_ar(i, extra=""):
    return f'SUMIFS({AR_RANGE["amt"]},{AR_RANGE["exp"]},">="&{WEEK_START[i]},{AR_RANGE["exp"]},"<="&{WEEK_END[i]}{extra})'

def sumifs_ap(i, extra=""):
    return f'SUMIFS({AP["amt"]},{AP["pd"]},">="&{WEEK_START[i]},{AP["pd"]},"<="&{WEEK_END[i]},{AP["m"]},"<>Card"{extra})'

def sumifs_rec(i):
    return f'SUMIFS({REC["amt"]},{REC["pd"]},">="&{WEEK_START[i]},{REC["pd"]},"<="&{WEEK_END[i]},{REC["m"]},"<>Card")'

def sumifs_pay(i, col):
    return f'SUMIFS(Payroll!${col}$5:${col}$6,Payroll!$C$5:$C$6,">="&{WEEK_START[i]},Payroll!$C$5:$C$6,"<="&{WEEK_END[i]})'

def in_week(i, dateref, amtref):
    return f'IF(AND({dateref}>={WEEK_START[i]},{dateref}<={WEEK_END[i]}),{amtref},0)'

r = 8
put(C, f"A{r}", "OPENING CASH (all bank accounts)", BOLD, fill=SUB)
OPEN_ROW = r
for col in "BCDEFGH": C[f"{col}{r}"].fill = SUB
r += 2
put(C, f"A{r}", "RECEIPTS", H2)
r += 1
receipt_rows = []
lines = [
    ("Customer receipts - invoices already overdue", lambda i: "=" + sumifs_ar(i, f',{AR_RANGE["status"]},"Overdue"'), "Receivables tab, status 'Overdue' (1-30 days late). Excludes the two 3-month-old debts (Meet Magic, Property NSW) which are dated past 1 Nov."),
    ("Customer receipts - invoices not yet due", lambda i: "=" + sumifs_ar(i, f',{AR_RANGE["status"]},"Current"'), "Receivables tab, status 'Current'. Includes Aware Super $225,887 (exp 9 Oct)."),
    ("September contract invoices (raised 30 Sep, paid early)", lambda i: "=" + sumifs_ar(i, f',{AR_RANGE["status"]},"To be raised"'), "Deloitte, CBA x4, Bankwest, Aware maintenance - based on how quickly each paid its August invoice. PwC pushed to November."),
    ("October-dated invoices paid within October", lambda i: f"={AS['quickpay']}/3" if i >= 2 else "=0", "Assumptions tab. Card/Stripe payers and small event invoices; spread over weeks 3-5."),
    ("Interest received", lambda i: f"={AS['interest']}" if i == 4 else "=0", "Assumptions tab. Savings interest lands at month end."),
]
for label, fn, note in lines:
    put(C, f"A{r}", label)
    for i, col in enumerate(week_cols()):
        put(C, f"{col}{r}", fn(i), BLACK, MONEY)
    put(C, f"G{r}", f"=SUM(B{r}:F{r})", BLACK, MONEY)
    put(C, f"H{r}", note, GREY, wrap=True)
    receipt_rows.append(r)
    r += 1
put(C, f"A{r}", "TOTAL RECEIPTS", BOLD, fill=TOTAL)
for col in "BCDEFG":
    put(C, f"{col}{r}", f"=SUM({col}{receipt_rows[0]}:{col}{receipt_rows[-1]})", BOLD, MONEY, fill=TOTAL)
C[f"H{r}"].fill = TOTAL
TOT_REC = r
r += 2
put(C, f"A{r}", "PAYMENTS", H2)
r += 1
put(C, f"A{r}", "Payroll & payroll-related", BOLD, fill=SUB)
for col in "BCDEFGH": C[f"{col}{r}"].fill = SUB
r += 1
pay_rows = []
payroll_lines = [
    ("Net wages - pay run 1 (Tue 6 Oct) and pay run 2 (Tue 20 Oct)", lambda i: "=" + sumifs_pay(i, "E"), "Payroll tab: gross x net %. ABA file uploaded to CommBiz on the Monday, paid Tuesday."),
    ("Superannuation (direct debit, same week as pay run)", lambda i: "=" + sumifs_pay(i, "G"), "Payroll tab: gross x super %. Employment Hero clearing house debits the cheque account (Checklist steps 69-70)."),
    ("Payroll tax NSW - September wages (due 7 Oct)", lambda i: "=" + in_week(i, AS["ptdate"], AS["ptnsw"]), "Assumptions tab; August actual used as proxy for September."),
    ("Payroll tax other states - September wages (due 7 Oct)", lambda i: "=" + in_week(i, AS["ptdate"], AS["ptoth"]), "Assumptions tab; August actual used as proxy."),
    ("PAYG withholding + GST - Q1 BAS (only if paid in Oct)", lambda i: "=" + in_week(i, AS["basdate"], f"{AS['bastoggle']}*({AS['baspayg']}+{AS['basgst']}+{AS['basinst']})"), "Assumptions tab toggle. Default 0 = lodged via tax agent, due 25 Nov. Set toggle to 1 to show it on 28 Oct."),
]
for label, fn, note in payroll_lines:
    put(C, f"A{r}", label)
    for i, col in enumerate(week_cols()):
        put(C, f"{col}{r}", fn(i), BLACK, MONEY)
    put(C, f"G{r}", f"=SUM(B{r}:F{r})", BLACK, MONEY)
    put(C, f"H{r}", note, GREY, wrap=True)
    pay_rows.append(r); r += 1
put(C, f"A{r}", "Subtotal payroll & payroll-related", BOLD)
for col in "BCDEFG":
    put(C, f"{col}{r}", f"=SUM({col}{pay_rows[0]}:{col}{pay_rows[-1]})", BOLD, MONEY)
    C[f"{col}{r}"].border = TOP
SUB_PAY = r
r += 2
put(C, f"A{r}", "Suppliers & overheads", BOLD, fill=SUB)
for col in "BCDEFGH": C[f"{col}{r}"].fill = SUB
r += 1
sup_rows = []
supplier_lines = [
    ("Complete AV Solutions - Aware seminar room (held until Aware pays)", lambda i: "=" + sumifs_ap(i, f',{AP["sup"]},"Complete AV Solutions"'), "Payables tab. $178,119 released the week after Aware's $225,887 arrives."),
    ("Rent - Investa (September rent, due 22 Oct)", lambda i: "=" + sumifs_ap(i, f',{AP["sup"]},"Investa Asset Management"'), "Payables tab. October rent is due ~22 Nov."),
    ("Other bills already in Xero (EFT/BPAY/DD)", lambda i: "=" + sumifs_ap(i, f',{AP["sup"]},"<>Complete AV Solutions",{AP["sup"]},"<>Investa Asset Management"'), "Payables tab, all other open bills by planned payment date. Card-paid bills excluded."),
    ("October recurring bills not yet in Xero (First Focus, RJW, advisory etc.)", lambda i: "=" + sumifs_rec(i), "Payables tab, second block - from Xero repeating-bill templates."),
    ("Credit card settlements (PL + DL cards)", lambda i: f"={AS['cc']}*0.2", "Assumptions tab, spread evenly. Covers every 'paid via CC' bill and card spend."),
    ("Other direct costs not yet billed (subcontractors, hires, equipment)", lambda i: f"={AS['directcost']}*0.2", "Assumptions tab estimate, spread evenly. Biggest swing factor after Aware/Complete AV."),
    ("Other overheads not itemised", lambda i: f"={AS['overhead']}*0.2", "Assumptions tab estimate, spread evenly."),
]
for label, fn, note in supplier_lines:
    put(C, f"A{r}", label)
    for i, col in enumerate(week_cols()):
        put(C, f"{col}{r}", fn(i), BLACK, MONEY)
    put(C, f"G{r}", f"=SUM(B{r}:F{r})", BLACK, MONEY)
    put(C, f"H{r}", note, GREY, wrap=True)
    sup_rows.append(r); r += 1
put(C, f"A{r}", "Subtotal suppliers & overheads", BOLD)
for col in "BCDEFG":
    put(C, f"{col}{r}", f"=SUM({col}{sup_rows[0]}:{col}{sup_rows[-1]})", BOLD, MONEY)
    C[f"{col}{r}"].border = TOP
SUB_SUP = r
r += 1
put(C, f"A{r}", "TOTAL PAYMENTS", BOLD, fill=TOTAL)
for col in "BCDEFG":
    put(C, f"{col}{r}", f"={col}{SUB_PAY}+{col}{SUB_SUP}", BOLD, MONEY, fill=TOTAL)
C[f"H{r}"].fill = TOTAL
TOT_PAY = r
r += 2
put(C, f"A{r}", "NET CASH FLOW FOR THE WEEK", BOLD)
for col in "BCDEFG":
    put(C, f"{col}{r}", f"={col}{TOT_REC}-{col}{TOT_PAY}", BOLD, MONEY)
NET = r
r += 1
put(C, f"A{r}", "CLOSING CASH (all bank accounts)", BOLD, fill=TOTAL)
for i, col in enumerate(week_cols()):
    put(C, f"{col}{r}", f"={col}{OPEN_ROW}+{col}{NET}", BOLD, MONEY, fill=TOTAL)
put(C, f"G{r}", f"=F{r}", BOLD, MONEY, fill=TOTAL)
C[f"H{r}"].fill = TOTAL
CLOSE = r
# opening cash links
put(C, f"B{OPEN_ROW}", f"={AS['opening']}", GREEN, MONEY, fill=SUB, bold=True)
for i in range(1, 5):
    col = get_column_letter(2 + i); prev = get_column_letter(1 + i)
    put(C, f"{col}{OPEN_ROW}", f"={prev}{CLOSE}", BLACK, MONEY, fill=SUB, bold=True)
put(C, f"G{OPEN_ROW}", f"=B{OPEN_ROW}", BLACK, MONEY, fill=SUB, bold=True)
put(C, f"H{OPEN_ROW}", "Xero cash position 28 Sep 2026. Total of every bank account, not just the cheque account.", GREY, wrap=True)
r += 2
put(C, f"A{r}", "MEMO", H2); r += 1
put(C, f"A{r}", "Total payroll-related cash out (net wages + super + payroll tax + BAS)")
for col in "BCDEFG": put(C, f"{col}{r}", f"={col}{SUB_PAY}", BLACK, MONEY)
r += 1
put(C, f"A{r}", "Cash needed in the cheque account on pay day (net wages + super)")
for i, col in enumerate(week_cols()):
    put(C, f"{col}{r}", f"={col}{pay_rows[0]}+{col}{pay_rows[1]}", BLACK, MONEY)
put(C, f"H{r}", "Checklist step 63-65: transfer from savings before upload, leaving ~$20k buffer.", GREY, wrap=True)
r += 1
put(C, f"A{r}", "Closing cash as weeks of payroll cover (net wages + super per fortnight / 2)")
for i, col in enumerate(week_cols()):
    put(C, f"{col}{r}", f"=IF(Payroll!$H$5=0,0,{col}{CLOSE}/(Payroll!$H$5/2))", BLACK, '0.0')
put(C, f"H{r}", "How many weeks of net payroll + super the closing balance would cover if no more cash came in.", GREY, wrap=True)
r += 1
put(C, f"A{r}", "Receivables still expected after 1 Nov (not in this forecast)")
put(C, f"G{r}", f'=SUMIFS({AR_RANGE["amt"]},{AR_RANGE["exp"]},">"&{FC_END})', BLACK, MONEY)
put(C, f"H{r}", "PwC September contract invoices, Meet Magic, Property NSW.", GREY, wrap=True)
r += 2
put(C, f"A{r}", "KEY RISKS & THINGS TO CONFIRM", H2); r += 1
risks = [
    "1. Aware Super $225,887 (INV-10596, due 9 Oct) is 39% of all expected receipts. Complete AV $178,119 must stay on hold until it lands - if Aware slips a week, move both lines (Receivables row 12 / Payables row 6).",
    "2. Ventia $58,025 is a month overdue and Ventia's other three invoices ($36,677) are assumed to come in 2-3 weeks late. Chase this week.",
    "3. Payroll is the fixed commitment: ~$100k leaves the cheque account on 6 Oct and again on 20 Oct (net pay + super). Payroll tax ~$9k on 7 Oct. Make sure the cheque account is topped up from savings before the ABA upload.",
    "4. BAS: the September-quarter BAS (Sep PAYG withholding ~$50k + net GST + any PAYG instalment) is assumed to be paid in November via the tax agent. If it is self-lodged it is due 28 Oct - flip the toggle on the Assumptions tab.",
    "5. 'Other direct costs not yet billed' ($100k) and 'Other overheads' ($15k) are estimates from the FY27 GL run-rate. October is budgeted as a heavy Production month, so hires and subcontractor spend could exceed this.",
    "6. Two debts totalling $24,809 (Meet Magic, Property NSW) are 3-4 months overdue and excluded. Escalate or provide for them.",
    "7. Not visible through the Xero connector and therefore NOT included: director/shareholder drawings, dividends, loan movements, insurance instalments, income tax instalment amounts, and the cheque/savings split of the opening balance.",
]
for t in risks:
    put(C, f"A{r}", t, BLACK, wrap=True)
    C.merge_cells(f"A{r}:H{r}")
    C.row_dimensions[r].height = 30
    r += 1
C.freeze_panes = "B7"
C.sheet_view.zoomScale = 90
for ws in wb.worksheets:
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
from openpyxl.workbook.properties import CalcProperties
wb.calculation = CalcProperties(fullCalcOnLoad=True)
wb.save(OUT)
print("saved", OUT)
