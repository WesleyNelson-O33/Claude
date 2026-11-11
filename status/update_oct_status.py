"""Refresh the October status workbook to where October stands on 11 November 2026."""
import copy, openpyxl
from openpyxl.styles import Font
SRC = "/root/.claude/uploads/fc902e4e-1220-55ba-999f-a6d0cad6db14/36b84039-October_status_and_Duncan_answer_-_revenue_change_111126_at_11.49am.xlsx"
OUT = "/home/user/Claude/status/out/October status and Duncan answer - updated 111126.xlsx"
wb = openpyxl.load_workbook(SRC); ws = wb["October"]
TEXT, BLUE, HEAD, NOTE = ws["A7"], ws["B7"], ws["A6"], ws["A2"]
def put(ref, value, like):
    c = ws[ref]; c.value = value; c._style = copy.copy(like._style); return c

ws["A1"] = "OCTOBER 2026 IS NEARLY CLOSED"
ws["A2"] = "Month end journals posted as at 11 November 2026, 10.46am: payroll tax, prepayment releases, leave, time in lieu and the expense accruals. Figures are from Xero at that time."
ws["A3"] = "Two things are not settled: the WIP journal, and the reversal of September's time in lieu accrual. Do not quote the month until both are done."
ws["B10"] = "=1051482.95"
ws["A12"] = "October is the highest month of the year on billed revenue. WIP is not confirmed, so revenue for the month can still move."
ws["A14"] = "October as it stands"
ws["B16"] = "=-533599.76"
ws["B20"] = -216445.24
ws["A20"] = "Overheads posted"
ws["A21"] = "Figure on screen"

# the journals table, rows 24 to 39
for r in range(24, 40):
    for c in "ABC": ws[f"{c}{r}"].value = None
put("A24", "October month end journals", HEAD); put("B24", "October amount", HEAD); put("C24", "Status", HEAD)
rows = [
    ("Work in progress journal", "not quantified", "Not confirmed. Ask before quoting the month."),
    ("Payroll tax, NSW, other states and the September top up", 8712.82, "Posted"),
    ("Prepayment releases, all accounts", 13477.78, "Posted"),
    ("   of which other insurance", 5161.63, "Posted"),
    ("   of which workers compensation", 874.34, "Posted"),
    ("   of which travel insurance", 73.33, "Posted"),
    ("Leave expense, annual", 4189.75, "Posted"),
    ("Leave expense, long service", 2217.12, "Posted"),
    ("Accounting fees", 3417.00, "Posted"),
    ("Business advisory", 1200.00, "Posted"),
    ("Instant asset write off accrual", 6425.91, "Posted"),
    ("Time in lieu accrual, October", 7288.96, "Posted"),
    ("September time in lieu accrual reversal", -4534.37, "Not posted. Costs are overstated by this until it is."),
]
for i, (a, b, c) in enumerate(rows, start=25):
    put(f"A{i}", a, TEXT); put(f"B{i}", b, BLUE if isinstance(b, (int, float)) else TEXT); put(f"C{i}", c, NOTE)
put("A38", "Figure on screen after the reversal still to come", TEXT); put("B38", "=B21-B37", ws["B21"])
put("A39", "An indication only. The WIP journal can move revenue either way. There is no reliable October net figure until WIP is confirmed and the time in lieu reversal is posted.", NOTE)

# checklist
for r in range(48, 53): ws[f"A{r}"].value = None; ws[f"B{r}"].value = None
for i, (a, b) in enumerate([("October payroll tax lodged", "Done, 9 November, 5,996.99"),
                             ("Payroll tax journals imported", "Done, all three"),
                             ("Prepayment releases run", "Done, 13,477.78"),
                             ("October WIP journal run", "Not confirmed"),
                             ("September time in lieu accrual reversed", "Not done, 4,534.37")], start=48):
    put(f"A{i}", a, TEXT); put(f"B{i}", b, NOTE)

# as billed, and where the movement is
ws["C70"] = -533599.76
for ref, v in (("C74", 374214.10), ("C75", -163301.13), ("C76", 228240.84), ("C77", -120549.90), ("C78", 149149.44), ("C79", 132231.00)):
    ws[ref] = v
ws["A81"] = "October 2025 column checked against Xero on 11 November 2026. It agrees."
ws["A81"]._style = copy.copy(NOTE._style)

q = wb["Q1 Integration"]
q["B29"] = -5050.10
q["A31"] = "October is still being finalised. The subcontract labour line moved from 13,255.00 to a credit of 5,050.10 on 11 November, so treat this as indicative only."
wb.save(OUT); print("saved")
