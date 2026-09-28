"""Build the CTS weekly cash flow forecast workbook (v2: forecast vs actual, P&L layout).

Data sources: Xero (cash position, aged receivables, aged payables, repeating
bills/invoices, cash-basis P&L, bill line items) pulled 28 Sep 2026; FY27 GL in
the Controller Pack; payroll manuals/transcripts; Complete AV meeting transcript.
"""
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.properties import CalcProperties

OUT = "/home/user/Claude/cashflow/CTS Cash Flow Forecast - Oct 2026.xlsx"
F = "Arial"
def font(size=10, bold=False, color=None, italic=False):
    return Font(name=F, size=size, bold=bold, color=color, italic=italic)
BLACK, BOLD, BLUE, GREEN = font(), font(bold=True), font(color="0000FF"), font(color="008000")
WHITE_B = font(bold=True, color="FFFFFF"); WHITE = font(color="FFFFFF")
TITLE = font(14, True); H2 = font(11, True); GREY = font(9, color="666666", italic=True)
NAVY = PatternFill("solid", fgColor="002060")
G1 = PatternFill("solid", fgColor="F2F2F2"); G2 = PatternFill("solid", fgColor="D9D9D9"); G3 = PatternFill("solid", fgColor="BFBFBF")
INPUT = PatternFill("solid", fgColor="FFF2CC"); YELLOW = PatternFill("solid", fgColor="FFFF00"); NOFILL = PatternFill(fill_type=None)
thin = Side(style="thin", color="7F7F7F"); TOP = Border(top=thin); TOPBOT = Border(top=thin, bottom=Side(style="double", color="7F7F7F"))
MONEY = '$#,##0;($#,##0);-'; MONEY2 = '$#,##0.00;($#,##0.00);-'; PCT = '0.0%'; DATE = 'ddd d mmm yy'; DATE_S = 'd mmm yy'
D = date

def put(ws, ref, val, f=BLACK, fmt=None, fill=None, wrap=False, align=None, border=None):
    c = ws[ref]; c.value = val; c.font = f
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if wrap: c.alignment = Alignment(wrap_text=True, vertical="top")
    if align: c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if border: c.border = border
    return c

def band(ws, row, c1, c2, fill, f=None):
    for col in range(c1, c2 + 1):
        c = ws.cell(row=row, column=col); c.fill = fill
        if f is not None: c.font = f

wb = openpyxl.Workbook()

# =============================================================== Assumptions
A = wb.active; A.title = "Assumptions"
A.column_dimensions["A"].width = 54; A.column_dimensions["B"].width = 16; A.column_dimensions["C"].width = 100
put(A, "A1", "CTS Cash Flow Forecast - Assumptions & Inputs", TITLE)
put(A, "A2", "Blue = input you can change. Black = formula. Bright yellow = key assumption to confirm. All $ AUD, GST inclusive (cash actually moving).", GREY)
rows = [
    ("TIMING & OPENING POSITION", None, None, None),
    ("Forecast start (Monday of this week)", D(2026, 9, 28), DATE, "Mon 28 Sep 2026. Five weeks: Wk1 28 Sep-4 Oct, Wk2 5-11 Oct, Wk3 12-18 Oct, Wk4 19-25 Oct, Wk5 26 Oct-1 Nov."),
    ("Opening bank balance - all accounts", 909159.55, MONEY2, "Xero cash position 28 Sep 2026 (cheque, savings and credit cards combined). Payroll is paid from the cheque account with top-ups from savings - check the cheque balance separately before each pay run."),
    ("Number of weeks", 5, None, "Cash Flow tab has Forecast and Actual columns for each week."),
    ("PAYROLL (fortnightly, paid Tuesday of processing week)", None, None, None),
    ("Gross wages per fortnight", 111000, MONEY, "Average of the last three known pay runs: 11 Aug $108,325, 25 Aug $115,340 (GL journals #69396, #69669) and the two September runs $217,062 combined (Xero cash P&L 1-28 Sep). See Payroll tab."),
    ("Net pay as % of gross", 0.77, PCT, "25 Aug run: gross ~$115k, ABA file ~$88-89k (Payroll Part 6 transcript, 03:16-04:12) = 76.8%. Balance is PAYG withholding and salary-sacrifice deductions."),
    ("PAYG withheld as % of gross", "=1-B10", PCT, "Formula: 1 - net %. Paid to the ATO via IAS/BAS, not on payday."),
    ("Superannuation as % of gross", 0.12, PCT, "Superannuation Guarantee 12%. 25 Aug run: SG $13,677 on gross $115,340 = 11.9%. Salary-sacrifice super (~$1,050 a run) is in the same super batch but comes out of gross pay, so it is not added here."),
    ("Employee reimbursements per pay run (expense claims)", 1000, MONEY, "PLACEHOLDER - not visible in Xero through the connector. Per diems ($75/night) are paid inside payroll and are already in net pay. Enter the typical expense-claim amount paid with each pay run."),
    ("Pay date - run 1 (fortnight ending Fri 2 Oct)", D(2026, 10, 6), DATE, "Pay date is always the Tuesday of the processing week (Payroll Part 2 transcript; Checklist step 13). Super is direct-debited the same week (Checklist step 69); Payday Super applies from 1 Jul 2026."),
    ("Pay date - run 2 (fortnight ending Fri 16 Oct)", D(2026, 10, 20), DATE, "Next run is Tue 3 Nov (fortnight ending 30 Oct) - outside this forecast."),
    ("Payroll tax NSW - September return (due 7 Oct)", 6295, MONEY, "August actual accrual $6,120.40 + $174.67 (GL journals #69671, #70351). NSW monthly return due the 7th of the following month. Replace with the lodged figure."),
    ("Payroll tax other states (ACT/QLD/SA/VIC/WA) - Sep (due 7 Oct)", 2700, MONEY, "August actual accrual $2,700 (GL #69670; July $3,000). All due on the 7th of the following month."),
    ("Payroll tax payment date", D(2026, 10, 7), DATE, "Wed 7 Oct 2026."),
    ("BAS / PAYG WITHHOLDING", None, None, None),
    ("Q1 BAS paid inside this forecast? (1 = yes, 0 = no)", 0, None, "Sep-quarter BAS is due 28 Oct if self-lodged, or 25 Nov through the tax agent. The Jun-quarter BAS was paid ~25 Aug (Part 6 transcript, 04:40), which matches the agent date, so default is 0 (paid November). Set to 1 to show it on 28 Oct."),
    ("Sep PAYG withholding (goes on the Q1 BAS)", "=B11*217062", MONEY, "Formula: PAYG % x September gross wages $217,062 (Xero cash P&L). July and August withholding were paid on the monthly IAS (21 Aug, 21 Sep), so no IAS falls in October."),
    ("Net GST payable for Sep quarter - ESTIMATE", 50000, MONEY, "Rough: GST on ~$1.5m quarterly receipts less GST on ~$1.0m purchases. Replace with the Xero Activity Statement figure."),
    ("PAYG income-tax instalment for Sep quarter", 0, MONEY, "Unknown - not visible through the Xero connector. Enter the ATO instalment if one applies."),
    ("BAS payment date if paid in October", D(2026, 10, 28), DATE, "Self-lodgement due date."),
    ("COSTS NOT YET BILLED - MONTHLY ESTIMATES (spread evenly over the 5 weeks)", None, None, None),
    ("Sub-contract labour", 40000, MONEY, "September cash P&L: Sub-Contract Labour INTEGRATION $31,990 + PRD $7,154. Bills for October work paid inside the month."),
    ("Equipment & service hires", 30000, MONEY, "September cash P&L: Service hires $17,344 + Equipment hires $11,427."),
    ("Project equipment purchases (not yet billed)", 20000, MONEY, "Beyond the Complete AV / Crestron / Fredon bills already listed. Budget flags October as a heavy Production month ($587k income), so this could run higher."),
    ("Direct travel & freight (client jobs)", 7000, MONEY, "September cash P&L: Direct Travel ONS $4,120 + PRD $1,572 + freight ~$1,300."),
    ("Client subscriptions & licences (not yet billed)", 3000, MONEY, "September cash P&L: Subscriptions & Licences expense $3,573 beyond the itemised Appspace/Spacera/Eptura bills."),
    ("Recruitment (Seek, LinkedIn, agency)", 5000, MONEY, "September cash P&L: Recruitment $11,742 (elevated by a casual concierge campaign). Half that assumed."),
    ("Accounting, legal & advisory (beyond the 8020 retainer)", 4000, MONEY, "Jun-Aug cash P&L: Accounting fees $7,733 and Legal $2,900 over three months."),
    ("Insurance", 3000, MONEY, "Jun-Aug cash P&L: Other insurance $5,073 + workers comp $2,336 over three months (~$2.5k/month)."),
    ("Office, amenities & entertainment", 2000, MONEY, "Jun-Aug cash P&L: staff amenities, office supplies, staff entertainment ~$4k/month; September ran lower."),
    ("Bank & merchant fees", 1000, MONEY, "September cash P&L: Bank charges $2,311 + Stripe fees $271 (Sep included a one-off)."),
    ("Credit card settlements per month (PL + DL cards)", 25000, MONEY, "September transfers from cheque to cards: $15,000 + $4,224.75 (PL) and $5,718.88 (DL) = $24,944 (Xero bank transactions). Covers every bill marked 'paid via CC', which are excluded from the Payables totals to avoid double counting."),
    ("RECEIPTS - ESTIMATES", None, None, None),
    ("October-dated invoices collected within October", 20000, MONEY, "September evidence: invoices dated and paid in Sep (Ubank $10,867, RSNSW $3,587, CSIRO $1,971, Garvan, Elkiem, Bankwest event, St Vincent's) = ~$20k. Spread over weeks 3-5."),
    ("Interest received on savings", 600, MONEY, "September $648.92 (Xero cash P&L). Lands at month end (week 5)."),
]
r = 4
for label, val, fmt, note in rows:
    if val is None and fmt is None:
        put(A, f"A{r}", label, WHITE_B, fill=NAVY); band(A, r, 2, 3, NAVY)
    else:
        put(A, f"A{r}", label)
        put(A, f"B{r}", val, BLACK if (isinstance(val, str) and val.startswith("=")) else BLUE, fmt)
        put(A, f"C{r}", note, GREY, wrap=True)
    r += 1
lab = {A[f"A{i}"].value: i for i in range(4, r)}
def AS(prefix):
    for k, v in lab.items():
        if k and k.startswith(prefix): return f"Assumptions!$B${v}"
    raise KeyError(prefix)
KEY = {"start": AS("Forecast start"), "opening": AS("Opening bank"), "gross": AS("Gross wages"), "netpct": AS("Net pay as %"),
       "paygpct": AS("PAYG withheld as %"), "superpct": AS("Superannuation as %"), "reimb": AS("Employee reimbursements"),
       "pay1": AS("Pay date - run 1"), "pay2": AS("Pay date - run 2"), "ptnsw": AS("Payroll tax NSW"), "ptoth": AS("Payroll tax other"),
       "ptdate": AS("Payroll tax payment date"), "bastoggle": AS("Q1 BAS paid"), "baspayg": AS("Sep PAYG withholding"),
       "basgst": AS("Net GST"), "basinst": AS("PAYG income-tax"), "basdate": AS("BAS payment date"),
       "est_sub": AS("Sub-contract labour"), "est_hire": AS("Equipment & service hires"), "est_equip": AS("Project equipment purchases"),
       "est_travel": AS("Direct travel & freight"), "est_subs": AS("Client subscriptions"), "est_recruit": AS("Recruitment ("),
       "est_prof": AS("Accounting, legal"), "est_ins": AS("Insurance"), "est_office": AS("Office, amenities"), "est_bank": AS("Bank & merchant"),
       "cc": AS("Credit card settlements"), "quickpay": AS("October-dated invoices"), "interest": AS("Interest received")}
assert A[KEY["paygpct"].split("$B$")[1] and "B" + KEY["paygpct"].split("$B$")[1]].value == "=1-B10", "PAYG % formula must reference net % (B10)"
assert lab["Net pay as % of gross"] == 10 and lab["PAYG withheld as % of gross"] == 11
for k in ("netpct", "superpct", "ptnsw", "ptoth", "bastoggle", "basgst", "est_sub", "est_hire", "est_equip", "reimb"):
    A[KEY[k].replace("Assumptions!", "").replace("$", "")].fill = YELLOW
r += 1
put(A, f"A{r}", "WEEK TABLE (calculated)", WHITE_B, fill=NAVY); band(A, r, 2, 3, NAVY); r += 1
put(A, f"A{r}", "Week", BOLD, fill=G1); put(A, f"B{r}", "Start (Mon)", BOLD, fill=G1); put(A, f"C{r}", "End (Sun)", BOLD, fill=G1); r += 1
WEEK_START, WEEK_END = [], []
for i in range(5):
    put(A, f"A{r}", i + 1); put(A, f"B{r}", f"={KEY['start']}+{7*i}", BLACK, DATE); put(A, f"C{r}", f"=B{r}+6", BLACK, DATE)
    WEEK_START.append(f"Assumptions!$B${r}"); WEEK_END.append(f"Assumptions!$C${r}"); r += 1
FC_END = WEEK_END[-1]

# =============================================================== Payroll
P = wb.create_sheet("Payroll")
for col, w in zip("ABCDEFGHIJ", (34, 16, 16, 14, 14, 14, 14, 14, 16, 60)): P.column_dimensions[col].width = w
put(P, "A1", "Payroll schedule - October 2026", TITLE)
put(P, "A2", "Fortnightly pay cycle. Pay date = Tuesday of the processing week. Super is direct-debited the same week. PAYG withholding is paid later via IAS/BAS.", GREY)
hdr = ["Pay run", "Fortnight ending", "Pay date", "Gross wages", "Net pay (ABA file)", "PAYG withheld", "Super (direct debit)", "Reimbursements", "Cash out on payday (net + super + reimb.)", "Notes"]
for j, h in enumerate(hdr): put(P, f"{L(j+1)}4", h, WHITE_B, fill=NAVY, wrap=True)
P.row_dimensions[4].height = 30
runs = [("Run 1 - Oct", D(2026, 10, 2), KEY["pay1"], "Timesheets close Fri 2 Oct, processed Mon 5 Oct, paid Tue 6 Oct."),
        ("Run 2 - Oct", D(2026, 10, 16), KEY["pay2"], "Last pay run of the month: accrue bonus leave (Checklist step 48) - no cash effect.")]
for i, (name, fe, payref, note) in enumerate(runs):
    rr = 5 + i
    put(P, f"A{rr}", name); put(P, f"B{rr}", fe, BLUE, DATE); put(P, f"C{rr}", f"={payref}", GREEN, DATE)
    put(P, f"D{rr}", f"={KEY['gross']}", GREEN, MONEY); put(P, f"E{rr}", f"=D{rr}*{KEY['netpct']}", BLACK, MONEY)
    put(P, f"F{rr}", f"=D{rr}*{KEY['paygpct']}", BLACK, MONEY); put(P, f"G{rr}", f"=D{rr}*{KEY['superpct']}", BLACK, MONEY)
    put(P, f"H{rr}", f"={KEY['reimb']}", GREEN, MONEY); put(P, f"I{rr}", f"=E{rr}+G{rr}+H{rr}", BLACK, MONEY); put(P, f"J{rr}", note, GREY, wrap=True)
put(P, "A7", "Total October pay runs", BOLD, fill=G2); band(P, 7, 2, 10, G2)
for col in "DEFGHI": put(P, f"{col}7", f"=SUM({col}5:{col}6)", BOLD, MONEY, fill=G2)
put(P, "A9", "Payroll tax for September wages (paid October)", H2)
put(P, "A10", "NSW - due 7 Oct"); put(P, "B10", f"={KEY['ptnsw']}", GREEN, MONEY); put(P, "C10", f"={KEY['ptdate']}", GREEN, DATE)
put(P, "A11", "Other states - due 7 Oct"); put(P, "B11", f"={KEY['ptoth']}", GREEN, MONEY); put(P, "C11", f"={KEY['ptdate']}", GREEN, DATE)
put(P, "A12", "Total payroll tax", BOLD, fill=G2); put(P, "B12", "=B10+B11", BOLD, MONEY, fill=G2); band(P, 12, 3, 3, G2)
put(P, "A13", "Payroll tax for October wages is due 7 Nov - outside this forecast.", GREY)
put(P, "A15", "Reference - actual pay runs this financial year", H2)
for j, h in enumerate(["Pay run", "Pay date", "Gross wages", "Super batch (SG + salary sacrifice)", "Batch % of gross", "Source"]):
    put(P, f"{L(j+1)}16", h, WHITE_B, fill=NAVY, wrap=True)
P.row_dimensions[16].height = 30
actuals = [("FE 10 Jul", D(2026, 7, 14), 104335.03, 13289.41, "Controller Pack GL_Paste journals #69527 / #69434"),
           ("FE 24 Jul", D(2026, 7, 28), 101030.18, 13068.87, "GL journals #68743 / #68744"),
           ("FE 7 Aug", D(2026, 8, 11), 108324.58, 13783.25, "GL journals #69396 / #69397"),
           ("FE 21 Aug", D(2026, 8, 25), 115339.80, 14677.93, "GL journals #69669 / #69663 (the run in the training videos)"),
           ("FE 4 Sep + FE 18 Sep (two runs)", D(2026, 9, 22), 217061.76, 27815.63, "Xero cash-basis P&L 1-28 Sep 2026: all Salaries & Wages accounts; Direct + Indirect Superannuation")]
for i, (n, d, g, s, src) in enumerate(actuals):
    rr = 17 + i
    put(P, f"A{rr}", n); put(P, f"B{rr}", d, BLUE, DATE); put(P, f"C{rr}", g, BLUE, MONEY); put(P, f"D{rr}", s, BLUE, MONEY)
    put(P, f"E{rr}", f"=IF(C{rr}=0,0,D{rr}/C{rr})", BLACK, PCT); put(P, f"F{rr}", src, GREY)
put(P, "A22", "Average gross per fortnight (last 4 fortnights)", BOLD, fill=G2); put(P, "B22", "=(C19+C20+C21)/4", BOLD, MONEY, fill=G2); put(P, "C22", "Aug runs plus the two Sep runs, over four fortnights", GREY)
put(P, "A23", "SG is 12%. The batch ratio above is higher because the GL super accounts include salary-sacrifice super (25 Aug: $13,677 SG + $1,052 salary sacrifice). The forecast uses 12%; the salary-sacrifice portion is deducted from gross pay so it is not double counted.", GREY)
put(P, "A25", "Reference - payroll tax accruals this financial year", H2)
for j, h in enumerate(["Month", "NSW", "Other states", "Total", "Source"]): put(P, f"{L(j+1)}26", h, WHITE_B, fill=NAVY)
put(P, "A27", "July 2026"); put(P, "B27", 5170.80, BLUE, MONEY); put(P, "C27", 3000, BLUE, MONEY); put(P, "D27", "=B27+C27", BLACK, MONEY); put(P, "E27", "GL journals #68926 (NSW) and #68927 (other states)", GREY)
put(P, "A28", "August 2026"); put(P, "B28", 6295.07, BLUE, MONEY); put(P, "C28", 2700, BLUE, MONEY); put(P, "D28", "=B28+C28", BLACK, MONEY); put(P, "E28", "GL journals #69671 + #70351 (NSW) and #69670 (other states)", GREY)
put(P, "A29", "FY27 budget has payroll tax at ~$7.4k/month; actuals are running ~$9k/month, so actuals are used.", GREY)

# =============================================================== Receivables
R = wb.create_sheet("Receivables")
for j, w in enumerate((34, 12, 44, 12, 12, 14, 13, 16, 11, 8, 62)): R.column_dimensions[L(j+1)].width = w
put(R, "A1", "Receivables - every open invoice and when the cash is expected", TITLE)
put(R, "A2", "Source: Xero aged receivables 28 Sep 2026 ($580,292.39, 38 invoices) plus September contract invoices still to be raised (Xero repeating templates). Expected date (blue) is a judgement from each customer's recent paying pattern - change it and the forecast moves.", GREY)
hdr = ["Customer", "Invoice", "Reference", "Invoice date", "Due date", "Amount (incl GST)", "Status", "Expected receipt date", "Days after due", "Week", "Basis for expected date"]
for j, h in enumerate(hdr): put(R, f"{L(j+1)}4", h, WHITE_B, fill=NAVY, wrap=True)
R.row_dimensions[4].height = 30
ar = [
 ("APT Management Services (APA)", "INV-10622", "PO-00030884", D(2026,9,10), D(2026,10,10), 4352.45, "Current", D(2026,10,9), "APA paid its last invoice the day after issue"),
 ("APT Management Services (APA)", "INV-10628", "APA RUOK Event 10.09.2026", D(2026,9,16), D(2026,10,16), 627.00, "Current", D(2026,10,16), "On due date"),
 ("ASX Operations", "INV-10599", "ASX Employee Town Hall 13.08.2026", D(2026,8,27), D(2026,9,26), 1661.00, "Overdue", D(2026,10,2), "ASX paid its Aug invoices ~4 days late"),
 ("ASX Operations", "INV-10624", "ASX CEO Connect Series 01.09.2026", D(2026,9,15), D(2026,10,15), 4591.40, "Current", D(2026,10,16), "Due date + a day"),
 ("ASX Operations", "INV-10623", "ASX Live 01.09.2026", D(2026,9,15), D(2026,10,15), 8544.80, "Current", D(2026,10,16), "Due date + a day"),
 ("Australian Payments Network", "INV-10574", "Logitech room licence renewal", D(2026,8,20), D(2026,9,19), 1092.30, "Overdue", D(2026,10,2), "Small, overdue - chase this week"),
 ("Aware Super", "INV-10583", "Annual National Maintenance (Aug)", D(2026,8,31), D(2026,9,30), 12100.00, "Current", D(2026,9,30), "Aware pays on or near due date"),
 ("Aware Super", "INV-10596", "Seminar Room AV Upgrade", D(2026,9,9), D(2026,10,9), 225887.40, "Current", D(2026,10,9), "KEY RECEIPT. Complete AV materials payment ($136,163) is released the week after this lands. If Aware slips, hold Complete AV."),
 ("Aware Super", "INV-10635", "Appspace licensing subscription (12 months)", D(2026,9,18), D(2026,10,18), 23317.11, "Current", D(2026,10,16), "Due Sun 18 Oct - expect Fri 16 Oct. Funds the Appspace bill ($20,985) due 24 Oct."),
 ("Aware Super", "INV-10646", "Perth office BYOD USB-C replacement", D(2026,9,24), D(2026,10,24), 3252.70, "Current", D(2026,10,23), "Due Sat 24 Oct - expect Fri 23 Oct"),
 ("Aware Super", "INV-10619", "Sydney 24.16 UC Engine remediation variation", D(2026,9,25), D(2026,10,25), 738.10, "Current", D(2026,10,23), "Due Sun 25 Oct - expect Fri 23 Oct"),
 ("Downer EDI", "INV-10587", "Full Year Results Briefing 20.08.2026", D(2026,8,25), D(2026,9,24), 12808.69, "Overdue", D(2026,10,2), "4 days overdue - assume paid this week"),
 ("Downer EDI", "INV-10611", "Full Year Employee Briefing 25.08.2026", D(2026,8,31), D(2026,9,30), 26855.13, "Current", D(2026,10,7), "Due 30 Sep; allow a week"),
 ("Generation-e Productivity Solutions", "INV-10532", "PAU26-407 (Jul)", D(2026,7,31), D(2026,8,30), 13627.68, "Overdue", D(2026,10,7), "A month overdue - chase. Assume week 2"),
 ("Generation-e Productivity Solutions", "INV-10573", "PAU26-407 (Aug)", D(2026,8,31), D(2026,9,30), 13627.68, "Current", D(2026,10,21), "Slow payer - assume 3 weeks after due"),
 ("ICC Sydney", "INV-10630", "Speaker presentation technicians - ICAS 2026", D(2026,9,16), D(2026,10,16), 1040.88, "Current", D(2026,10,16), "On due date"),
 ("ICC Sydney", "INV-10633", "Speaker presentation technicians - ISE", D(2026,9,16), D(2026,10,16), 5131.51, "Current", D(2026,10,16), "On due date"),
 ("Jaazaniah Salanoa (staff)", "INV-10621", "Goget hire recovery", D(2026,8,31), D(2026,9,30), 66.64, "Current", D(2026,10,6), "Staff recovery - likely via payroll"),
 ("Meet Magic", "INV-10464", "Human AI Event highlight videos 29.05.2026", D(2026,6,19), D(2026,6,26), 5075.07, "Overdue 3 months", D(2026,11,30), "DOUBTFUL - 3 months overdue. Excluded from October. Escalate."),
 ("Microsoft", "INV-10554", "PO 10160645", D(2026,7,31), D(2026,8,30), 13845.70, "Overdue", D(2026,10,7), "A month overdue - usually a PO/portal issue. Assume week 2"),
 ("ON24", "INV-10552", "PO #25962 (Jul)", D(2026,7,31), D(2026,8,30), 2832.50, "Overdue", D(2026,10,7), "A month overdue - chase"),
 ("ON24", "INV-10589", "PO25962 (Aug)", D(2026,8,26), D(2026,9,25), 3005.75, "Overdue", D(2026,10,7), "Assume paid with the July one"),
 ("Property NSW", "INV-10396", "Penrith BDA Condeco desk booking (PO20285637)", D(2026,5,6), D(2026,6,5), 19733.45, "Overdue 3 months", D(2026,11,30), "DOUBTFUL - almost 4 months overdue. Excluded from October. Escalate."),
 ("PwC Services", "INV-10535", "PO2601702104 (Jul)", D(2026,7,31), D(2026,8,30), 13097.70, "Overdue", D(2026,10,7), "A month overdue. PwC paid INV-10577 on its due date, so these look like PO problems - chase"),
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
assert abs(sum(x[5] for x in ar) - 580292.39) < 0.01
to_raise = [
 ("Deloitte Services Trust", "TBR", "CONTRACT - September technicians", D(2026,9,30), D(2026,10,30), 113987.99, "To be raised", D(2026,10,14), "Aug invoice (31 Aug) paid 14 Sep - Deloitte pays ~2 weeks after issue. Repeating template $113,987.99/month"),
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
rr = 5; ar_first = rr
for cust, inv, ref, idate, due, amt, status, exp, basis in ar + to_raise:
    put(R, f"A{rr}", cust); put(R, f"B{rr}", inv); put(R, f"C{rr}", ref)
    put(R, f"D{rr}", idate, BLUE, DATE_S); put(R, f"E{rr}", due, BLUE, DATE_S); put(R, f"F{rr}", amt, BLUE, MONEY2); put(R, f"G{rr}", status, BLUE)
    put(R, f"H{rr}", exp, BLUE, DATE_S, fill=YELLOW if amt > 50000 else None)
    put(R, f"I{rr}", f"=H{rr}-E{rr}", BLACK, '0')
    put(R, f"J{rr}", f'=IF(H{rr}<{KEY["start"]},"Before",IF(H{rr}>{FC_END},"After",INT((H{rr}-{KEY["start"]})/7)+1))', BLACK)
    put(R, f"K{rr}", basis, GREY, wrap=True); rr += 1
ar_last = rr - 1
put(R, f"A{rr}", "TOTAL", BOLD, fill=G2); band(R, rr, 2, 11, G2); put(R, f"F{rr}", f"=SUM(F{ar_first}:F{ar_last})", BOLD, MONEY2, fill=G2)
put(R, f"A{rr+1}", "of which open in Xero at 28 Sep (should equal $580,292.39)", GREY); put(R, f"F{rr+1}", f'=SUMIFS(F{ar_first}:F{ar_last},G{ar_first}:G{ar_last},"<>To be raised")', BLACK, MONEY2)
put(R, f"A{rr+2}", "of which expected inside the forecast window", GREY); put(R, f"F{rr+2}", f'=SUMIFS(F{ar_first}:F{ar_last},H{ar_first}:H{ar_last},">="&{KEY["start"]},H{ar_first}:H{ar_last},"<="&{FC_END})', BLACK, MONEY2)
put(R, f"A{rr+3}", "of which expected after 1 Nov (PwC Sep invoices, Meet Magic, Property NSW)", GREY); put(R, f"F{rr+3}", f"=F{rr}-F{rr+2}", BLACK, MONEY2)
R.freeze_panes = "A5"; R.auto_filter.ref = f"A4:K{ar_last}"
AR = dict(amt=f"Receivables!$F${ar_first}:$F${ar_last}", exp=f"Receivables!$H${ar_first}:$H${ar_last}")

# =============================================================== Payables
Y = wb.create_sheet("Payables")
for j, w in enumerate((32, 48, 12, 14, 10, 24, 16, 8, 14, 62)): Y.column_dimensions[L(j+1)].width = w
put(Y, "A1", "Payables - every open bill, plus October's recurring bills not yet in Xero", TITLE)
put(Y, "A2", "Source: Xero aged payables 28 Sep 2026 ($283,042.53, 24 bills; Complete AV split per the 28 Sep meeting) and Xero repeating-bill templates. 'Card' rows are paid on the PL/DL credit cards and are covered by the credit-card settlement line on the Cash Flow tab, so they are excluded from the category totals. The P&L category drives which Cash Flow line each bill lands on.", GREY, wrap=True)
Y.row_dimensions[2].height = 40
hdr = ["Supplier", "Bill / reference", "Due date", "Amount (incl GST)", "Method", "P&L category", "Planned payment date", "Week", "Source", "Notes"]
for j, h in enumerate(hdr): put(Y, f"{L(j+1)}4", h, WHITE_B, fill=NAVY, wrap=True)
Y.row_dimensions[4].height = 30
X = "Xero bill"; T = "Repeating template"
ap = [
 ("Appspace, Inc.", "INV00133660 - Appspace licensing for Aware Super", D(2026,10,24), 20985.43, "EFT", "Client subscriptions", D(2026,10,23), X, "Pay once Aware's matching invoice INV-10635 ($23,317) is received (expected 16 Oct)"),
 ("Complete AV Solutions", "INV41565 - Aware seminar room: MATERIALS portion (hardware, cables, freight, waste, net of the $4,848.68 claim-balance reduction)", D(2026,9,26), 136163.28, "EFT", "Equipment purchases", D(2026,10,14), X, "Agreed 28 Sep meeting (Graham/Jordan): pay the materials part now, hold labour until Aware signs off. Materials $128,633.48 less $4,848.68 = $123,784.80 ex GST. Pay the week after Aware's $225,887 lands (exp 9 Oct)."),
 ("Complete AV Solutions", "INV41565 - Aware seminar room: LABOUR & PROGRAMMING portion - HELD", D(2026,9,26), 41955.98, "EFT", "Equipment purchases", D(2026,11,30), X, "HELD until Aware sign-off and handover docs (marked-up drawings, training docs). $38,141.80 ex GST = Labour & Programming $23,320 + Labour $7,196 + Labour $6,561.80 + Lighting programming $1,064. Dated after the forecast so it is excluded."),
 ("Crestron ANZ", "INV950936412 - Aware SYD 28.06 equipment", D(2026,10,19), 7103.26, "EFT", "Equipment purchases", D(2026,10,19), X, "On due date"),
 ("Employsure", "41 of 60", D(2026,9,6), 880.00, "EFT", "Professional fees", D(2026,10,1), X, "Overdue - pay this week"),
 ("Energy Australia", "INV260382065288 - energy 23 Jul-22 Aug", D(2026,9,29), 377.13, "Card", "Electricity", D(2026,9,29), X, "Direct debit via PL credit card"),
 ("Energy Australia", "INV260412412804 - electricity 23 Aug-21 Sep", D(2026,10,14), 696.25, "Card", "Electricity", D(2026,10,14), X, "Direct debit via PL credit card"),
 ("Eptura Australia", "INV-85230 - Proxyclick renewal APA VMS yr 2", D(2026,10,1), 3627.04, "EFT", "Client subscriptions", D(2026,10,1), X, "On due date"),
 ("Expedia", "ITN73550041732636 - accommodation D Hurley", D(2026,9,22), 808.95, "Card", "Direct travel & freight", D(2026,9,22), X, "Earlier Expedia bookings went on the PL card"),
 ("Fredon Technology", "INV11S-26080021 - DTTL 8PSQ Dream Big microphones", D(2026,9,23), 24708.20, "EFT", "Equipment purchases", D(2026,10,7), X, "Overdue; bill note 'DH to confirm'. Assume released week 2"),
 ("Investa Asset Management", "INV5310034149 - Rent September 2026 (incl. pest)", D(2026,10,22), 29280.72, "EFT", "Rent", D(2026,10,22), X, "On due date. October rent will be invoiced ~22 Oct and due ~22 Nov"),
 ("Mailchimp", "MC27286583 - September", D(2026,9,18), 37.62, "Card", "Dues & subscriptions", D(2026,9,18), X, "PL card"),
 ("Optus", "INV000594614144 - office internet & landline Sep", D(2026,9,23), 629.66, "Card", "Telephone & internet", D(2026,9,23), X, "PL card"),
 ("Optus", "INV000594999543 - staff mobiles Sep", D(2026,10,12), 804.11, "BPAY", "Telephone & internet", D(2026,10,12), X, "BPAY on due date"),
 ("Persona Health", "INV-18315 - fitness for duty assessment", D(2026,9,10), 996.60, "EFT", "Other", D(2026,10,21), X, "Bill note: GC advised hold until he confirms. Assumed released week 4"),
 ("Posh Services (Ambius)", "INV21736844 - indoor plants Sep", D(2026,10,9), 245.00, "EFT", "Office", D(2026,10,9), X, "On due date"),
 ("Ricoh Australia", "INV15759052 - printing Aug", D(2026,9,30), 235.60, "DD", "Office", D(2026,9,30), X, "Direct debit"),
 ("RJW Group Holdings", "INV-0048 - website design balance", D(2026,9,30), 6600.00, "EFT", "Advertising", D(2026,10,2), X, "Bill note: check with DL before payment"),
 ("Seek", "INV702112646 - job ad casual AV concierge", D(2026,9,18), 550.00, "Card", "Recruitment", D(2026,9,18), X, "PL card"),
 ("SPACERA", "INV-0056 - Aware Super national support (Sep)", D(2026,10,1), 5728.80, "EFT", "Client subscriptions", D(2026,10,1), X, "On due date"),
 ("Tec Art", "INV136603 - PRD asset purchase", D(2026,9,22), 657.50, "Card", "Equipment purchases", D(2026,9,22), X, "PL card"),
 ("Tec Art", "INV136603CR - refund", D(2026,9,24), -657.50, "Card", "Equipment purchases", D(2026,9,24), X, "Refund to PL card"),
 ("Transport for NSW", "Opal top-up T Wood", D(2026,9,18), 20.00, "Card", "Direct travel & freight", D(2026,9,18), X, "Card"),
 ("Virgin Australia", "SYMUMZ-3 - rescheduled flights D Hurley", D(2026,9,18), 501.91, "Card", "Direct travel & freight", D(2026,9,18), X, "Earlier Virgin bookings went on the PL card"),
 ("Xero", "INV-56302114 - subscription Sep", D(2026,9,24), 107.00, "Card", "Dues & subscriptions", D(2026,9,24), X, "PL card"),
 ("First Focus IT", "CORE Intelligent Managed Services - October", D(2026,10,16), 13755.77, "DD", "IT support", D(2026,10,16), T, "Repeating bill template, direct debit"),
 ("First Focus IT", "Managed Services Azure - October", D(2026,10,16), 710.50, "DD", "IT support", D(2026,10,16), T, "Repeating bill template, direct debit"),
 ("RJW Group Holdings", "Digital growth marketing - October", D(2026,10,21), 4675.00, "EFT", "Advertising", D(2026,10,21), T, "Repeating bill template"),
 ("Cleveland Partners (8020 Advisors)", "Monthly advisory fees", D(2026,10,7), 3850.00, "EFT", "Professional fees", D(2026,10,7), T, "Template $3,850/month (Sep cash P&L shows $5,000 business advisory). Date assumed"),
 ("SPACERA", "Aware Super national support - October", D(2026,10,31), 5728.80, "EFT", "Client subscriptions", D(2026,10,31), T, "Repeating bill, next due 31 Oct"),
 ("LiveU", "SOLO standard data pack - October", D(2026,10,14), 1485.00, "EFT", "Client subscriptions", D(2026,10,14), T, "Repeating bill, due 14 Oct"),
 ("Employsure", "42 of 60 - October", D(2026,10,5), 880.00, "EFT", "Professional fees", D(2026,10,5), T, "Repeating bill, due 5 Oct"),
 ("The CEO Institute", "Syndicate membership (DL) - quarterly", D(2026,10,8), 2310.00, "EFT", "Dues & subscriptions", D(2026,10,8), T, "Repeating bill every 3 months, due 8 Oct"),
 ("Telstra", "Nighthawk data plans - October", D(2026,10,31), 130.00, "EFT", "Telephone & internet", D(2026,10,31), T, "Repeating bill"),
 ("Employment Hero", "Subscription - September", D(2026,10,30), 2048.20, "Card", "Dues & subscriptions", D(2026,9,30), T, "PL card - in credit card settlement line"),
 ("Goget Carshare", "September usage", D(2026,10,30), 1783.07, "Card", "Direct travel & freight", D(2026,9,30), T, "PL card - in credit card settlement line"),
 ("Zoom", "Monthly subscription", D(2026,10,28), 323.97, "Card", "Dues & subscriptions", D(2026,10,28), T, "PL card"),
 ("Insphire (Current RMS)", "October", D(2026,10,16), 431.00, "Card", "Dues & subscriptions", D(2026,10,16), T, "PL card"),
 ("Optus", "Office internet & landline - October", D(2026,10,21), 629.66, "Card", "Telephone & internet", D(2026,10,21), T, "PL card"),
 ("Adobe / Spotify / Avangate / Telstra recharges / misc", "Small monthly subscriptions", D(2026,10,15), 700.00, "Card", "Dues & subscriptions", D(2026,10,15), T, "PL/DL cards - approx"),
 ("Investa Asset Management", "Rent October 2026", D(2026,11,22), 29280.72, "EFT", "Rent", D(2026,11,22), T, "Invoiced ~22 Oct, due ~22 Nov - AFTER this forecast"),
 ("Optus", "Staff mobiles - October", D(2026,11,8), 804.06, "BPAY", "Telephone & internet", D(2026,11,8), T, "Due 8 Nov - AFTER this forecast"),
]
assert abs(sum(x[3] for x in ap if x[7] == X) - 283042.53) < 0.02, sum(x[3] for x in ap if x[7] == X)
rr = 5; ap_first = rr
for sup, ref, due, amt, method, cat, pdate, src, note in ap:
    put(Y, f"A{rr}", sup); put(Y, f"B{rr}", ref, wrap=True); put(Y, f"C{rr}", due, BLUE, DATE_S)
    put(Y, f"D{rr}", amt, BLUE, MONEY2); put(Y, f"E{rr}", method, BLUE); put(Y, f"F{rr}", cat, BLUE)
    put(Y, f"G{rr}", pdate, BLUE, DATE_S, fill=YELLOW if amt > 50000 else None)
    put(Y, f"H{rr}", f'=IF(G{rr}<{KEY["start"]},"Before",IF(G{rr}>{FC_END},"After",INT((G{rr}-{KEY["start"]})/7)+1))', BLACK)
    put(Y, f"I{rr}", src); put(Y, f"J{rr}", note, GREY, wrap=True); rr += 1
ap_last = rr - 1
put(Y, f"A{rr}", "TOTAL open bills in Xero (should equal $283,042.53)", BOLD, fill=G2); band(Y, rr, 2, 10, G2)
put(Y, f"D{rr}", f'=SUMIFS(D{ap_first}:D{ap_last},I{ap_first}:I{ap_last},"{X}")', BOLD, MONEY2, fill=G2)
put(Y, f"A{rr+1}", "of which paid by card (excluded - see credit card line)", GREY); put(Y, f"D{rr+1}", f'=SUMIFS(D{ap_first}:D{ap_last},E{ap_first}:E{ap_last},"Card",I{ap_first}:I{ap_last},"{X}")', BLACK, MONEY2)
put(Y, f"A{rr+2}", "Recurring templates (non-card) falling inside the forecast", GREY); put(Y, f"D{rr+2}", f'=SUMIFS(D{ap_first}:D{ap_last},I{ap_first}:I{ap_last},"{T}",E{ap_first}:E{ap_last},"<>Card",G{ap_first}:G{ap_last},">="&{KEY["start"]},G{ap_first}:G{ap_last},"<="&{FC_END})', BLACK, MONEY2)
put(Y, f"A{rr+3}", "Non-card bills dated after 1 Nov (Complete AV labour hold, Oct rent, Oct mobiles)", GREY); put(Y, f"D{rr+3}", f'=SUMIFS(D{ap_first}:D{ap_last},E{ap_first}:E{ap_last},"<>Card",G{ap_first}:G{ap_last},">"&{FC_END})', BLACK, MONEY2)
Y.freeze_panes = "A5"; Y.auto_filter.ref = f"A4:J{ap_last}"
AP = dict(amt=f"Payables!$D${ap_first}:$D${ap_last}", pd=f"Payables!$G${ap_first}:$G${ap_last}", m=f"Payables!$E${ap_first}:$E${ap_last}", cat=f"Payables!$F${ap_first}:$F${ap_last}")

# =============================================================== Cash Flow
C = wb.create_sheet("Cash Flow", 0)
C.column_dimensions["A"].width = 46
NW = 5
FCOL = [L(2 + 2*i) for i in range(NW)]      # forecast columns B, D, F, H, J
ACOL = [L(3 + 2*i) for i in range(NW)]      # actual columns   C, E, G, I, K
TF, TA, VAR, NOTE = L(2 + 2*NW), L(3 + 2*NW), L(4 + 2*NW), L(5 + 2*NW)   # L, M, N, O
for col in FCOL + ACOL + [TF, TA, VAR]: C.column_dimensions[col].width = 13
C.column_dimensions[NOTE].width = 64
LASTCOL = 5 + 2*NW
put(C, "A1", "Corporate Technology Services Pty Ltd", TITLE)
put(C, "A2", "Weekly Cash Flow Forecast vs Actual - 28 September to 1 November 2026", H2)
put(C, "A3", "How to use: forecast columns are formulas driven by the Receivables, Payables, Payroll and Assumptions tabs. Each week, type the actual cash movements into the yellow Actual column, enter the closing bank balance from CommBiz, and set 'Actuals entered' to 1. The next week's opening balance then rebases to the real bank balance, and the difference row shows anything still unexplained.", GREY, wrap=True)
C.merge_cells(f"A3:{NOTE}3"); C.row_dimensions[3].height = 42
# header block rows 5-8
for i in range(NW):
    f, a = FCOL[i], ACOL[i]
    put(C, f"{f}5", f"Week {i+1}", WHITE_B, fill=NAVY, align="center"); C[f"{a}5"].fill = NAVY; C.merge_cells(f"{f}5:{a}5")
    put(C, f"{f}6", f'=TEXT({WEEK_START[i]},"d mmm")&" - "&TEXT({WEEK_END[i]},"d mmm")', WHITE, fill=NAVY, align="center"); C[f"{a}6"].fill = NAVY; C.merge_cells(f"{f}6:{a}6")
    put(C, f"{f}7", "Forecast", WHITE_B, fill=NAVY, align="center"); put(C, f"{a}7", "Actual", WHITE_B, fill=NAVY, align="center")
put(C, "A5", "", fill=NAVY); put(C, "A6", "", fill=NAVY); put(C, "A7", "$ AUD (GST inclusive)", WHITE_B, fill=NAVY)
put(C, f"{TF}5", "5-week total", WHITE_B, fill=NAVY, align="center"); C[f"{TA}5"].fill = NAVY; C.merge_cells(f"{TF}5:{TA}5")
put(C, f"{TF}6", "", fill=NAVY); put(C, f"{TA}6", "", fill=NAVY)
put(C, f"{TF}7", "Forecast", WHITE_B, fill=NAVY, align="center"); put(C, f"{TA}7", "Actual", WHITE_B, fill=NAVY, align="center")
put(C, f"{VAR}5", "Variance", WHITE_B, fill=NAVY, align="center", wrap=True); put(C, f"{VAR}6", "(weeks with", WHITE, fill=NAVY, align="center"); put(C, f"{VAR}7", "actuals)", WHITE, fill=NAVY, align="center")
put(C, f"{NOTE}5", "Basis / source", WHITE_B, fill=NAVY); put(C, f"{NOTE}6", "", fill=NAVY); put(C, f"{NOTE}7", "", fill=NAVY)
FLAG = 8
put(C, f"A{FLAG}", "Actuals entered for this week? (1 = yes)", BOLD, fill=G1)
for i in range(NW):
    put(C, f"{FCOL[i]}{FLAG}", "", fill=G1); put(C, f"{ACOL[i]}{FLAG}", 0, BLUE, '0', fill=INPUT, align="center")
band(C, FLAG, 2 + 2*NW, LASTCOL, G1)
put(C, f"{NOTE}{FLAG}", "Set to 1 once the week's actuals and bank balance are in. Drives the opening balance of the following week and the variance column.", GREY, wrap=True)
FLAGS = [f"{ACOL[i]}${FLAG}" for i in range(NW)]

def wk_ar(i, extra=""):
    return f'SUMIFS({AR["amt"]},{AR["exp"]},">="&{WEEK_START[i]},{AR["exp"]},"<="&{WEEK_END[i]}{extra})'
def wk_ap(i, cats):
    parts = [f'SUMIFS({AP["amt"]},{AP["pd"]},">="&{WEEK_START[i]},{AP["pd"]},"<="&{WEEK_END[i]},{AP["m"]},"<>Card",{AP["cat"]},"{c}")' for c in cats]
    return "+".join(parts)
def wk_pay(i, col):
    return f'SUMIFS(Payroll!${col}$5:${col}$6,Payroll!$C$5:$C$6,">="&{WEEK_START[i]},Payroll!$C$5:$C$6,"<="&{WEEK_END[i]})'
def in_week(i, dateref, amtref):
    return f'IF(AND({dateref}>={WEEK_START[i]},{dateref}<={WEEK_END[i]}),{amtref},0)'
def est(key):
    return lambda i: f"={KEY[key]}/{NW}"

row = FLAG + 2
OPEN = row
put(C, f"A{row}", "OPENING BANK BALANCE", BOLD, fill=G2); band(C, row, 2, LASTCOL, G2)
put(C, f"{NOTE}{row}", "Week 1 = Xero cash position 28 Sep (all accounts). Later weeks = prior week's actual bank balance once actuals are entered, otherwise the prior forecast closing.", GREY, wrap=True)
row += 2

def section(title):
    global row
    put(C, f"A{row}", title, WHITE_B, fill=NAVY); band(C, row, 2, LASTCOL, NAVY); row += 1
def subhead(title):
    global row
    put(C, f"A{row}", title, BOLD); row += 1
def line(label, fn, note, actual_default=None):
    """fn(i) -> forecast formula for week i."""
    global row
    put(C, f"A{row}", "   " + label)
    for i in range(NW):
        put(C, f"{FCOL[i]}{row}", fn(i), BLACK, MONEY)
        put(C, f"{ACOL[i]}{row}", actual_default, BLUE, MONEY, fill=INPUT)
    put(C, f"{TF}{row}", f"=SUM({','.join(f'{c}{row}' for c in FCOL)})", BLACK, MONEY)
    put(C, f"{TA}{row}", f"=SUM({','.join(f'{c}{row}' for c in ACOL)})", BLACK, MONEY)
    put(C, f"{VAR}{row}", "=" + "+".join(f"({ACOL[i]}{row}-{FCOL[i]}{row})*{FLAGS[i]}" for i in range(NW)), BLACK, MONEY)
    put(C, f"{NOTE}{row}", note, GREY, wrap=True)
    row += 1; return row - 1
def total(label, rows_, fill=G1, bold=True, border=TOP):
    global row
    put(C, f"A{row}", label, BOLD, fill=fill); band(C, row, 2, LASTCOL, fill)
    for c in FCOL + ACOL + [TF, TA, VAR]:
        put(C, f"{c}{row}", "=" + "+".join(f"{c}{r_}" for r_ in rows_), BOLD, MONEY, fill=fill, border=border)
    row += 1; return row - 1

# ---- INCOME
section("INCOME")
inc_rows = []
inc_rows.append(line("Trading income - customer receipts (invoices)", lambda i: "=" + wk_ar(i) + (f"+{KEY['quickpay']}/3" if i >= 2 else ""), "Receivables tab by expected receipt date (open invoices, overdue invoices and the September contract invoices to be raised), plus the October quick-pay allowance from the Assumptions tab in weeks 3-5."))
inc_rows.append(line("Other income - interest received", lambda i: f"={KEY['interest']}" if i == NW-1 else "=0", "Assumptions tab. Savings interest at month end."))
TOT_INC = total("TOTAL INCOME", inc_rows, fill=G2)
row += 1
# ---- EXPENSES
section("OPERATING EXPENSES")
subhead("Employment")
emp = []
emp.append(line("Salaries & wages (net pay)", lambda i: "=" + wk_pay(i, "E"), "Payroll tab: gross x net %. Pay runs Tue 6 Oct and Tue 20 Oct (ABA file via CommBiz)."))
emp.append(line("Superannuation", lambda i: "=" + wk_pay(i, "G"), "Payroll tab: gross x 12%. Direct-debited by Employment Hero in the pay-run week."))
emp.append(line("PAYG withholding - monthly IAS", lambda i: "=0", "No IAS falls in October: July and August withholding were paid 21 Aug and 21 Sep; September's goes on the Q1 BAS (see BAS line)."))
emp.append(line("Payroll tax (NSW + other states)", lambda i: "=" + in_week(i, KEY["ptdate"], f"{KEY['ptnsw']}+{KEY['ptoth']}"), "Assumptions tab. September wages, due 7 Oct. August actuals used as proxy ($6,295 NSW + $2,700 other states)."))
emp.append(line("Employee reimbursements (expense claims)", lambda i: "=" + wk_pay(i, "H"), "Assumptions tab placeholder per pay run - replace with the real expense-claim run."))
SUB_EMP = total("Subtotal employment", emp)
row += 1
subhead("Cost of sales (client jobs)")
cos = []
cos.append(line("Project equipment & installation", lambda i: "=" + wk_ap(i, ["Equipment purchases"]) + "+" + est("est_equip")(i)[1:], "Payables tab (Complete AV materials $136,163 wk 3, Fredon $24,708 wk 2, Crestron $7,103 wk 4) plus the unbilled estimate from the Assumptions tab. Complete AV labour $41,956 is held past 1 Nov."))
cos.append(line("Sub-contract labour", est("est_sub"), "Assumptions tab estimate (Sep cash P&L ~$39k), spread evenly."))
cos.append(line("Equipment & service hires", est("est_hire"), "Assumptions tab estimate (Sep cash P&L ~$29k), spread evenly."))
cos.append(line("Subscriptions & licences (client)", lambda i: "=" + wk_ap(i, ["Client subscriptions"]) + "+" + est("est_subs")(i)[1:], "Payables tab (Appspace, Spacera, Eptura, LiveU) plus unbilled estimate."))
cos.append(line("Direct travel & freight", lambda i: "=" + wk_ap(i, ["Direct travel & freight"]) + "+" + est("est_travel")(i)[1:], "Payables tab (non-card) plus Assumptions estimate. Card-paid flights/accommodation sit in the credit card line."))
SUB_COS = total("Subtotal cost of sales", cos)
row += 1
subhead("Overheads")
ovh = []
ovh.append(line("Rent & outgoings", lambda i: "=" + wk_ap(i, ["Rent"]), "Payables tab: Investa September rent $29,281 due 22 Oct. October rent falls in November."))
ovh.append(line("IT network service & support", lambda i: "=" + wk_ap(i, ["IT support"]), "Payables tab: First Focus CORE + Azure direct debit ~16 Oct."))
ovh.append(line("Advertising & marketing", lambda i: "=" + wk_ap(i, ["Advertising"]), "Payables tab: RJW website balance $6,600 (wk 1) and RJW digital marketing $4,675 (wk 4)."))
ovh.append(line("Accounting, legal & advisory", lambda i: "=" + wk_ap(i, ["Professional fees"]) + "+" + est("est_prof")(i)[1:], "Payables tab (8020 Advisors retainer, Employsure) plus Assumptions estimate for accounting/legal."))
ovh.append(line("Dues & subscriptions", lambda i: "=" + wk_ap(i, ["Dues & subscriptions"]), "Payables tab non-card items (CEO Institute quarterly). Card subscriptions sit in the credit card line."))
ovh.append(line("Telephone & internet", lambda i: "=" + wk_ap(i, ["Telephone & internet"]), "Payables tab: Optus mobiles BPAY, Telstra Nighthawk. Optus internet is on the card."))
ovh.append(line("Recruitment", lambda i: "=" + wk_ap(i, ["Recruitment"]) + "+" + est("est_recruit")(i)[1:], "Assumptions estimate; Seek ads are on the card."))
ovh.append(line("Insurance", est("est_ins"), "Assumptions estimate (~$2.5k/month across business insurance and workers comp)."))
ovh.append(line("Office, printing & amenities", lambda i: "=" + wk_ap(i, ["Office"]) + "+" + est("est_office")(i)[1:], "Payables tab (Ricoh printing DD, Posh plants) plus Assumptions estimate."))
ovh.append(line("Electricity", lambda i: "=" + wk_ap(i, ["Electricity"]), "Energy Australia is direct-debited to the PL card, so it sits in the credit card line (shows nil here)."))
ovh.append(line("Bank & merchant fees", est("est_bank"), "Assumptions estimate (bank charges, Stripe fees)."))
ovh.append(line("Credit card settlements (card spend, all categories)", lambda i: f"={KEY['cc']}/{NW}", "Assumptions tab: ~$25k/month transferred to the PL and DL cards. Covers every 'Card' row on the Payables tab (subscriptions, travel, electricity, Seek, Goget, Employment Hero)."))
ovh.append(line("Other expenses", lambda i: "=" + wk_ap(i, ["Other"]), "Payables tab: Persona Health (on hold until GC confirms)."))
SUB_OVH = total("Subtotal overheads", ovh)
row += 1
subhead("Tax")
tax = []
tax.append(line("BAS - GST, PAYG withholding & instalment (Q1)", lambda i: "=" + in_week(i, KEY["basdate"], f"{KEY['bastoggle']}*({KEY['baspayg']}+{KEY['basgst']}+{KEY['basinst']})"), "Assumptions toggle. Default 0 = lodged via tax agent, due 25 Nov. Set to 1 to show ~$100k on 28 Oct."))
SUB_TAX = total("Subtotal tax", tax)
row += 1
TOT_EXP = total("TOTAL OPERATING EXPENSES", [SUB_EMP, SUB_COS, SUB_OVH, SUB_TAX], fill=G2)
row += 1
NET = row
put(C, f"A{row}", "NET CASH FLOW", BOLD, fill=G2); band(C, row, 2, LASTCOL, G2)
for c in FCOL + ACOL + [TF, TA, VAR]:
    put(C, f"{c}{row}", f"={c}{TOT_INC}-{c}{TOT_EXP}", BOLD, MONEY, fill=G2, border=TOP)
row += 1
CLOSE = row
put(C, f"A{row}", "CLOSING BANK BALANCE", BOLD, fill=G3); band(C, row, 2, LASTCOL, G3)
for i in range(NW):
    put(C, f"{FCOL[i]}{row}", f"={FCOL[i]}{OPEN}+{FCOL[i]}{NET}", BOLD, MONEY, fill=G3, border=TOPBOT)
    put(C, f"{ACOL[i]}{row}", f'=IF({FLAGS[i]}=1,{ACOL[i]}{OPEN}+{ACOL[i]}{NET},"")', BOLD, MONEY, fill=G3, border=TOPBOT)
put(C, f"{TF}{row}", f"={FCOL[-1]}{row}", BOLD, MONEY, fill=G3, border=TOPBOT)
last_actual = '""'
for i in range(NW):
    last_actual = f'IF({FLAGS[i]}=1,{ACOL[i]}{row},{last_actual})'
put(C, f"{TA}{row}", "=" + last_actual, BOLD, MONEY, fill=G3, border=TOPBOT)
put(C, f"{NOTE}{row}", "Opening + net cash flow. Actual column fills in once the week's flag is set to 1; the 5-week Actual total shows the latest completed week.", GREY, wrap=True)
row += 1
BANK = row
put(C, f"A{row}", "Actual bank balance per CommBiz (enter)", BOLD, fill=G1); band(C, row, 2, LASTCOL, G1)
for i in range(NW):
    put(C, f"{FCOL[i]}{row}", "", fill=G1); put(C, f"{ACOL[i]}{row}", None, BLUE, MONEY, fill=INPUT)
put(C, f"{NOTE}{row}", "Type the real closing balance (all accounts) at the end of each week.", GREY)
row += 1
DIFF = row
put(C, f"A{row}", "Unexplained difference (actual closing less bank)", BOLD, fill=G1); band(C, row, 2, LASTCOL, G1)
for i in range(NW):
    put(C, f"{FCOL[i]}{row}", "", fill=G1)
    put(C, f"{ACOL[i]}{row}", f'=IF({FLAGS[i]}=1,{ACOL[i]}{CLOSE}-{ACOL[i]}{BANK},"")', BOLD, MONEY, fill=G1)
put(C, f"{NOTE}{row}", "Should be nil once every movement is entered. A balance here means a receipt or payment is missing from the Actual column.", GREY, wrap=True)
# opening balance formulas (now that CLOSE/BANK known)
put(C, f"{FCOL[0]}{OPEN}", f"={KEY['opening']}", font(bold=True, color="008000"), MONEY, fill=G2)
put(C, f"{ACOL[0]}{OPEN}", f"={FCOL[0]}{OPEN}", BOLD, MONEY, fill=G2)
for i in range(1, NW):
    prevF, prevA = FCOL[i-1], ACOL[i-1]
    put(C, f"{FCOL[i]}{OPEN}", f"=IF({FLAGS[i-1]}=1,{prevA}{BANK},{prevF}{CLOSE})", BOLD, MONEY, fill=G2)
    put(C, f"{ACOL[i]}{OPEN}", f"={FCOL[i]}{OPEN}", BOLD, MONEY, fill=G2)
put(C, f"{TF}{OPEN}", f"={FCOL[0]}{OPEN}", BOLD, MONEY, fill=G2); put(C, f"{TA}{OPEN}", f"={ACOL[0]}{OPEN}", BOLD, MONEY, fill=G2)
# Memo
row += 2
put(C, f"A{row}", "MEMO", WHITE_B, fill=NAVY); band(C, row, 2, LASTCOL, NAVY); row += 1
put(C, f"A{row}", "   Cash needed in the cheque account on pay day (net + super + reimb.)")
for i in range(NW): put(C, f"{FCOL[i]}{row}", f"={FCOL[i]}{emp[0]}+{FCOL[i]}{emp[1]}+{FCOL[i]}{emp[4]}", BLACK, MONEY)
put(C, f"{NOTE}{row}", "Checklist steps 63-65: transfer from savings before the ABA upload, leaving ~$20k buffer.", GREY, wrap=True); row += 1
put(C, f"A{row}", "   Closing balance as weeks of payroll cover")
for i in range(NW): put(C, f"{FCOL[i]}{row}", f"=IF(Payroll!$I$5=0,0,{FCOL[i]}{CLOSE}/(Payroll!$I$5/2))", BLACK, '0.0')
put(C, f"{NOTE}{row}", "Weeks of net payroll + super the closing balance would cover with no further receipts.", GREY, wrap=True); row += 1
put(C, f"A{row}", "   Receipts expected after 1 Nov (not in this forecast)")
put(C, f"{TF}{row}", f'=SUMIFS({AR["amt"]},{AR["exp"]},">"&{FC_END})', BLACK, MONEY)
put(C, f"{NOTE}{row}", "PwC September contract invoices, Meet Magic, Property NSW.", GREY); row += 1
put(C, f"A{row}", "   Payments held past 1 Nov (not in this forecast)")
put(C, f"{TF}{row}", f'=SUMIFS({AP["amt"]},{AP["m"]},"<>Card",{AP["pd"]},">"&{FC_END})', BLACK, MONEY)
put(C, f"{NOTE}{row}", "Complete AV labour & programming $41,956 (held for Aware sign-off), October rent, October mobiles.", GREY, wrap=True); row += 2
put(C, f"A{row}", "KEY RISKS & THINGS TO CONFIRM", WHITE_B, fill=NAVY); band(C, row, 2, LASTCOL, NAVY); row += 1
risks = [
    "1. Aware Super $225,887 (INV-10596, due 9 Oct) is the single largest receipt. The Complete AV materials payment of $136,163 is timed for the week after it lands; if Aware slips, move both (Receivables row 12, Payables row 6).",
    "2. Complete AV labour and programming ($41,956 incl GST) is held until Aware signs off and handover documents arrive, per the 28 Sep meeting. Jordan flagged Complete AV may push to include programming or sub-trade costs - if agreed, add them to the Payables tab.",
    "3. Ventia $58,025 is a month overdue and its other three invoices ($36,677) are assumed 2-3 weeks late. PwC, Gen-e and Microsoft each have a month-old unpaid invoice. Chase this week.",
    "4. Payroll is the fixed commitment: about $100k leaves the cheque account on 6 Oct and again on 20 Oct, plus ~$9k payroll tax on 7 Oct. Top up the cheque account from savings before each ABA upload.",
    "5. BAS: the Sep-quarter BAS (Sep PAYG ~$50k + net GST + any instalment) is assumed paid 25 Nov via the tax agent. If self-lodged it is due 28 Oct - flip the toggle on the Assumptions tab.",
    "6. Unbilled cost estimates ($100k cost of sales, $15k overheads, $25k card) come from the FY27 GL run-rate. October is budgeted as a heavy Production month, so hires and sub-contractors could exceed this.",
    "7. Not visible through the Xero connector and NOT included: director drawings, dividends, loan movements, insurance instalment timing, income-tax instalment amounts, expense-claim runs, and the cheque/savings split of the opening balance.",
]
for t in risks:
    put(C, f"A{row}", t, BLACK, wrap=True); C.merge_cells(f"A{row}:{NOTE}{row}"); C.row_dimensions[row].height = 30; row += 1
C.freeze_panes = f"B{FLAG+1}"; C.sheet_view.zoomScale = 85; C.sheet_view.showGridLines = False
for i in range(NW):
    C.column_dimensions[ACOL[i]].width = 13
for ws in wb.worksheets:
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
wb.calculation = CalcProperties(fullCalcOnLoad=True)
wb.save(OUT)
print("saved", OUT, "rows", row)
