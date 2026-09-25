"""Fills a copy of the v4 tracker with a worked example for review.

One department row = one invoice line. Finance types Xero Invoice No, Date and Ex GST
on the matching Finance row. Job numbers, job names, cost centres and GL codes are real
(Xero job list on the v3 Lists sheet); every amount, invoice number and date is MADE UP.

Usage: python build_example.py <blank_v4.xlsx> <out.xlsx> [xero_figures.json]
"""
import datetime as dt
import json
import sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

SRC, OUT = sys.argv[1], sys.argv[2]
XERO = json.load(open(sys.argv[3])) if len(sys.argv) > 3 else {}
D = lambda y, m, d: dt.datetime(y, m, d)
wb = openpyxl.load_workbook(SRC)

# Each line: department cells, then the Xero invoice Finance types (or None)
ONSITE = [  # job, client, description, CC, type, tax, expected, PO, notes, invoice
    ("1144", "Bankwest", "Perth half-time tech - July", "ONSITE", "Contract", "GST 10%", 4000, "BW-PO-4471", "Monthly in arrears",
     ("INV-10101", D(2026, 7, 31), 4000)),
    ("1144", "Bankwest", "Perth half-time tech - August", "ONSITE", "Contract", "GST 10%", 4000, "BW-PO-4471", "",
     ("INV-10245", D(2026, 8, 31), 4000)),
    ("1144", "Bankwest", "Perth half-time tech - September", "ONSITE", "Contract", "GST 10%", 4000, "BW-PO-4471", "",
     ("INV-10390", D(2026, 9, 24), 4000)),
    ("1144", "Bankwest", "Perth half-time tech - October", "ONSITE", "Contract", "GST 10%", 4000, "BW-PO-4471",
     "Due end of October", None),
    ("1350", "CBA", "Brisbane onsite support - July", "ONSITE", "Contract", "GST 10%", 3750, "CBA-88213", "",
     ("INV-10102", D(2026, 7, 31), 3750)),
    ("1350", "CBA", "Brisbane onsite support - August", "ONSITE", "Contract", "GST 10%", 3750, "CBA-88213", "",
     ("INV-10246", D(2026, 8, 31), 3750)),
    ("20250731", "PWC", "National pool of hours - FY26 block", "ONSITE", "Ad-hoc", "GST 10%", 18000, "PWC-NPH-25",
     "Older job - invoiced June 2026", ("INV-9187", D(2026, 6, 15), 18000)),
    ("20250731", "PWC", "National pool of hours - August top-up", "ONSITE", "Ad-hoc", "GST 10%", 6000, "PWC-NPH-25", "",
     ("INV-10230", D(2026, 8, 14), 6000)),
    ("2026021101", "AusPayNet", "Annual AV support Jul-26 to Jun-27 (in advance)", "ONSITE", "Contract", "GST 10%", 12000,
     "APN-2026-Q", "Defer over 12 months", ("INV-10088", D(2026, 7, 1), 12000)),
    ("1530", "Deloitte (DTTL)", "Brisbane onsite tech - August cover", "ONSITE", "Ad-hoc", "GST 10%", 5500, "",
     "Work done in August, not invoiced yet", None),
    ("2192", "Bankwest", "Additional tech - August event", "ONSITE", "Ad-hoc", "GST 10%", 1800, "BW-PO-4502", "",
     ("INV-10231", D(2026, 8, 15), 1800)),
    ("2192", "Bankwest", "Remote support Aug-Oct (3 months)", "ONSITE", "Ad-hoc", "GST 10%", 1200, "BW-PO-4502",
     "Defer over 3 months", ("INV-10232", D(2026, 8, 15), 1200)),
    ("2193", "PWC", "Pool of Hours", "ONSITE", "Ad-hoc", "GST 10%", 5000, "No PO", "Added by reviewer", None),
]
PRODUCTION = [  # job, client, description, CC, type, tax, expected, cross hire, labour, event date, notes, invoice
    ("26073110", "Garvan Institute", "Weizmann dinner 20.8.26 - 40% deposit", "PRODUCTION", "Project", "GST 10%", 6000,
     2500, 3200, D(2026, 8, 20), "Costs for the whole job on this row", ("INV-10205", D(2026, 8, 5), 6000)),
    ("26073110", "Garvan Institute", "Weizmann dinner 20.8.26 - 60% balance", "PRODUCTION", "Project", "GST 10%", 9000,
     0, 0, D(2026, 8, 20), "", ("INV-10301", D(2026, 9, 2), 9000)),
    ("26073110V", "Garvan Institute", "Weizmann dinner 20.8.26 - videography", "VIDEO", "Project", "GST 10%", 3000,
     0, 900, D(2026, 8, 20), "Quoted $3,000", ("INV-10302", D(2026, 9, 2), 3500)),
    ("26073111", "Downer", "FYR & employee webcast 25.8.26", "PRODUCTION", "Project", "GST 10%", 9800, 1200, 2100,
     D(2026, 8, 25), "", ("INV-10270", D(2026, 8, 27), 9800)),
    ("26091601", "Automic", "September all hands 25.9.26", "PRODUCTION", "Project", "GST 10%", 7200, 800, 1500,
     D(2026, 9, 25), "Invoice after the event", None),
    ("2501738", "APA", "HYR25 & FYR25 results webcasts", "PRODUCTION", "Project", "GST 10%", 18500, 0, 0,
     D(2025, 8, 20), "Older job - invoiced Aug 2025", ("INV-9004", D(2025, 8, 20), 18500)),
    ("52202021", "Downer", "Pool of hours", "PRODUCTION", "Project", "GST 10%", 20000, 0, 4000, D(2026, 8, 31),
     "Added by reviewer", None),
    ("52202021V", "Downer", "Pool of hours - video", "VIDEO", "Project", "GST 10%", 3020000, 0, 11000, D(2026, 8, 31),
     "Added by reviewer - CHECK this value", None),
]
CONSULTING = [  # job, client, desc, CC, type, tax, expected, T lab, U equip, V subs, X ext lab, Y equip, Z subs, notes, invoice
    ("24121001", "Bankwest", "Perth studio - progress claim 1", "INTEGRATION", "Progress Claim", "GST 10%", 30000,
     0, 30000, 0, 4000, 24000, 0, "Older job - claimed May 2026", ("INV-9350", D(2026, 5, 15), 30000)),
    ("24121001", "Bankwest", "Perth studio - final claim", "INTEGRATION", "Progress Claim", "GST 10%", 15000,
     12000, 3000, 0, 0, 0, 0, "", ("INV-10120", D(2026, 7, 22), 15000)),
    ("25041502", "NSW Property (PDNSW)", "Desk booking licence Aug-26 to Jul-27", "INTEGRATION", "Subscription",
     "GST 10%", 9600, 0, 0, 9600, 0, 0, 7200, "Defer over 12 months", ("INV-10210", D(2026, 8, 3), 9600)),
    ("2722", "PWC", "Consulting retainer - July", "CONSULTING", "Contract", "GST 10%", 2500, 2500, 0, 0, 0, 0, 0, "",
     ("INV-10103", D(2026, 7, 31), 2500)),
    ("2722", "PWC", "Consulting retainer - August", "CONSULTING", "Contract", "GST 10%", 2500, 2500, 0, 0, 0, 0, 0, "",
     ("INV-10247", D(2026, 8, 31), 2500)),
    ("2722", "PWC", "Consulting retainer - September", "CONSULTING", "Contract", "GST 10%", 2500, 2500, 0, 0, 0, 0, 0,
     "", ("INV-10391", D(2026, 9, 24), 2500)),
    ("3040", "Corrs Chambers Westgarth", "Melbourne relocation - milestone 1", "CONSULTING", "Milestone", "GST 10%",
     6000, 6000, 0, 0, 3000, 0, 0, "", ("INV-10312", D(2026, 9, 3), None)),   # deliberate: no Ex GST typed
    ("3099", "Test Client", "EXAMPLE ERROR - job not in Xero and no tax code", "CONSULTING", "Project", None, 2000,
     2000, 0, 0, 0, 0, 0, "Deliberate mistake to show the checks", None),
]


def put(ws, row, **cells):
    for col, v in cells.items():
        if v not in (None, ""):
            ws[f"{col}{row}"] = v


fi = wb["Finance"]
H = {c.value: c.column_letter for c in fi[6]}
rowof = {fi[f"A{r}"].value: r for r in range(7, fi.max_row + 1)}


def invoice(rid, inv):
    if not inv:
        return
    r = rowof[rid]
    no, date, amt = inv
    fi[f"{H['Xero Invoice No']}{r}"] = no
    fi[f"{H['Xero Invoice Date']}{r}"] = date
    if amt is not None:
        fi[f"{H['Xero Invoiced Ex GST']}{r}"] = amt


on, pr, co = wb["Onsite"], wb["Production"], wb["Consulting"]
for i, (j, cl, ds, cc, it, tx, ex, po, nt, inv) in enumerate(ONSITE):
    put(on, 5 + i, C=j, F=cl, G=ds, H=cc, I=it, J=tx, O=ex, R=po, W=nt, V="Onsite Manager")
    invoice(f"ONS-{i + 1:04d}", inv)
for i, (j, cl, ds, cc, it, tx, ex, xh, lab, ed, nt, inv) in enumerate(PRODUCTION):
    put(pr, 5 + i, C=j, F=cl, G=ds, H=cc, I=it, J=tx, O=ex, AF=xh, AI=lab, V=ed, AT=nt, S=cl, AB="N")
    invoice(f"PRD-{i + 1:04d}", inv)
for i, (j, cl, ds, cc, it, tx, ex, t, u, v, x, y, z, nt, inv) in enumerate(CONSULTING):
    put(co, 5 + i, C=j, F=cl, G=ds, H=cc, I=it, J=tx, O=ex, T=t, U=u, V=v, X=x, Y=y, Z=z, AD=nt)
    invoice(f"CON-{i + 1:04d}", inv)

# --------------------------------------------------------------- Deferrals
df = wb["Deferrals"]
DH = {c.value: c.column_letter for c in df[6]}


def defer(row, **k):
    for key, v in k.items():
        df[f"{DH[key]}{row}"] = v


defer(7, **{"Type": "Revenue", "Invoice or Bill No": "INV-10088", "Invoice or Bill Date": D(2026, 7, 1),
            "Defer End": D(2027, 6, 30), "Notes": "AusPayNet annual support in advance"})
defer(8, **{"Type": "Revenue", "Invoice or Bill No": "INV-10210", "Invoice or Bill Date": D(2026, 8, 3),
            "Defer End": D(2027, 7, 31), "P&L GL Override": "41175",
            "Notes": "PDNSW 12-month licence - GL override: INTEGRATION maps to 42300, licences belong in 41175"})
defer(9, **{"Type": "Revenue", "Invoice or Bill No": "INV-10232", "Invoice or Bill Date": D(2026, 8, 15),
            "Defer End": D(2026, 10, 31), "Notes": "Bankwest remote support Aug-Oct"})
defer(10, **{"Type": "Cost", "Invoice or Bill No": "BILL-5521", "Invoice or Bill Date": D(2026, 8, 12),
             "Defer End": D(2026, 12, 31), "Job Number Override": "1144", "Amount Override": 1800,
             "P&L GL Override": "45010", "Notes": "Qantas flights & accommodation prepaid for the Perth contract"})

# ----------------------------------------------------------- WIP Movements
wp = wb["WIP Movements"]
put(wp, 5, A=D(2026, 8, 31), B="1530", E="ONSITE", G="August onsite cover worked, invoice not raised",
    H="Accrual - unbilled work", I="Revenue", J=2750, Q="N")
put(wp, 6, A=D(2026, 9, 30), B="1530", E="ONSITE", G="Reverse August accrual", H="Reversal of prior accrual",
    I="Revenue", J=-2750, Q="N")

# --------------------------------------------------------------- Work Won
ww = wb["Work Won"]
put(ww, 5, A=D(2026, 8, 31), B="Current RMS", C="RMS-7781", D="Garvan Institute", E="26073110", G="PRODUCTION",
    H="Weizmann dinner", I=15000, J="Won", K=D(2026, 7, 28), L=D(2026, 8, 31))
put(ww, 6, A=D(2026, 8, 31), B="Qwilr", C="Q-2291", D="Automic", E="26091601", G="PRODUCTION",
    H="September all hands", I=7200, J="Won", K=D(2026, 8, 18), L=D(2026, 9, 30))

# ------------------------------------------------------- Month-End (Aug-26)
me = wb["Month-End"]
me["C4"] = D(2026, 8, 31)
wb["Deferral Journal"]["C4"] = D(2026, 8, 31)
for cell, v in XERO.items():
    me[cell] = v

# ------------------------------------------------------------ Example Guide
PRIMARY, ACCENT, TINT, LINE = "1F3864", "BF8F00", "D9E2F3", "C9D3E3"
g = wb.create_sheet("Example Guide", 0)
for col, w in zip("ABCDE", (5, 30, 58, 72, 2)):
    g.column_dimensions[col].width = w
for c in range(1, 6):
    for r in (1, 2, 3):
        g.cell(r, c).fill = PatternFill("solid", fgColor=PRIMARY)
g["B2"] = "WORKED EXAMPLE  -  for review only, not live data"
g["B2"].font = Font(name="Arial", size=20, bold=True, color="FFFFFF")
g.row_dimensions[2].height = 34
g["B3"] = "Real job numbers from your Xero job list. Every amount, invoice number and date is made up."
g["B3"].font = Font(name="Arial", size=11, italic=True, color=TINT)
g.row_dimensions[3].height = 22
g["B5"] = ("ONE ROW = ONE INVOICE. Departments add one row per thing they will invoice. On Finance, each of those rows "
           "has three cream cells (columns G, H, I) where Finance types the Xero Invoice No, Date and Ex GST. "
           "Month-End and the Deferral Journal are set to August 2026.")
g["B5"].font = Font(name="Arial", size=11, bold=True, color="262626")
g["B5"].alignment = Alignment(wrap_text=True, vertical="center")
g.merge_cells("B5:D5")
g.row_dimensions[5].height = 48
heads = ["#", "Job / where to look", "What was entered", "What you should see"]
for j, h in enumerate(heads):
    c = g.cell(7, 1 + j, h)
    c.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=ACCENT if j == 3 else PRIMARY)
    c.alignment = Alignment(vertical="center", horizontal="center" if j == 0 else None, indent=0 if j == 0 else 1)
g.row_dimensions[7].height = 26
SCEN = [
    ("Finance sheet - how to type an invoice",
     "Finance row ONS-0001 (job 1144, 'Perth half-time tech - July', expected $4,000). Finance typed G: INV-10101, H: 31-Jul-26, I: 4,000.",
     "Variance $0, Compare 'Agrees'. Onsite sheet row 1 shows Invoice Date 31-Jul-26, Invoice No INV-10101, Revenue Ex GST $4,000, Invoiced 'Agrees'."),
    ("1144 Bankwest - monthly contract",
     "Onsite has four rows for 1144 (Jul, Aug, Sep, Oct). Finance typed Jul, Aug and Sep invoices.",
     "Jul-Sep rows 'Agrees'. October row 'Not invoiced - to invoice' $4,000 - the reminder to bill it. Lines on Job = 4. Finance: Job Total Expected $16,000, Job Total Invoiced $12,000."),
    ("1350 CBA", "Two rows (Jul, Aug), both invoiced.", "Both 'Agrees'."),
    ("20250731 PWC - older job",
     "Row 1 is the FY26 block invoiced 15-Jun-2026 ($18,000). Row 2 is an August top-up ($6,000).",
     "Both 'Agrees'. FY Summary counts only the August $6,000 in FY27 - the June 2026 invoice is outside the year."),
    ("2026021101 AusPayNet - revenue deferral",
     "Invoiced $12,000 on 01-Jul-26. Deferrals row 1: INV-10088, end Jun-27.",
     "Deferrals finds the job and $12,000 from the invoice number. Schedule: Jul -11,000, then +1,000 a month. Deferral Journal (Aug): Dr 11300 / Cr 41100 $1,000."),
    ("1530 DTTL - not invoiced + WIP accrual", "Expected $5,500, no invoice. WIP Movements: +$2,750 Aug, -$2,750 Sep.",
     "'Not invoiced - to invoice'. Month-End section 1 ONSITE WIP Movement +$2,750."),
    ("2192 Bankwest - two invoices same day",
     "Two rows: event $1,800 (INV-10231) and 3-month support $1,200 (INV-10232). Only INV-10232 is deferred.",
     "Each invoice has its own row, so Deferrals finds the $1,200 by itself - nothing to override. Schedule: Aug -800, Sep +400, Oct +400."),
    ("BILL-5521 - cost deferral (job 1144)", "Deferrals row 4: Cost, $1,800 bill, end Dec-26, job 1144, GL 45010.",
     "Schedule: Aug -1,440, then +360 Sep-Dec. Journal (Aug): Dr 11300 / Cr 45010 $1,440 with project 1144, Onsite."),
    ("26073110 Garvan - deposit + balance",
     "Production has two rows: 40% deposit $6,000 (INV-10205, Aug) and 60% balance $9,000 (INV-10301, Sep).",
     "Both 'Agrees'. Job total $15,000 on Finance. Costs sit on the deposit row: margin there $6,000 - $5,700."),
    ("26073110V Garvan video - over-invoiced", "Expected $3,000, Xero $3,500.",
     "'Xero higher - check dept', variance -$500."),
    ("26073111 Downer", "Expected $9,800, Xero $9,800.", "Agrees."),
    ("26091601 Automic", "Expected $7,200, no invoice.", "'Not invoiced - to invoice'. Work Won: won $7,200, $0 invoiced."),
    ("2501738 APA - older job", "Invoiced 20-Aug-2025.", "Agrees. Not in FY27 revenue."),
    ("24121001 BW studio - progress claims", "Claim 1 $30,000 (May 2026), final claim $15,000 (Jul 2026).",
     "Both 'Agrees'. FY Summary counts only the July $15,000 in FY27."),
    ("25041502 PDNSW licence - revenue deferral", "Invoiced $9,600 Aug. Deferrals row 2 with P&L GL Override 41175.",
     "Schedule: Aug -8,800, then +800 a month. Month-End INTEGRATION: invoiced $9,600, deferral -$8,800, recognised $800."),
    ("2722 PWC retainer", "Three monthly rows, all invoiced.", "All 'Agrees'."),
    ("3040 Corrs - DELIBERATE MISTAKE", "Finance typed INV-10312 and the date but forgot the Ex GST.",
     "Finance Issue: 'Invoice incomplete'. Compare 'Not invoiced - to invoice'. Month-End section 4 counts it."),
    ("3099 - DELIBERATE MISTAKE", "Job not in the Xero list, no tax code.",
     "Consulting 'Not Yet on Finance' = Fix: Tax Code. Finance Issue lists both. Month-End section 4 counts them."),
    ("Your added jobs: 2193, 52202021, 52202021V", "Kept as you typed them.",
     "All three show on Finance, flagged 'Job number not in the Xero job list'. 52202021V expects $3,020,000 - probably a typo."),
    ("Month-End (Aug-26)", "Xero revenue typed to match, except CONSULTING typed $250 low.",
     "Every cost centre Reconciled except CONSULTING: 'CHECK - $250.00'."),
    ("Deferral Journal (Aug-26)", "Nothing typed.",
     "Four lines (two revenue deferred, one released, one cost deferred) with project number, name, department, debit and credit accounts. Balanced."),
]
for k, (a, b, c) in enumerate(SCEN):
    r = 8 + k
    mistake = "MISTAKE" in a
    band = PatternFill("solid", fgColor="FCE4D6" if mistake else ("F4F7FB" if k % 2 else "FFFFFF"))
    for j, v in enumerate([k + 1, a, b, c]):
        cell = g.cell(r, 1 + j, v)
        cell.font = Font(name="Arial", size=10, bold=(j in (0, 1)), color=PRIMARY if j in (0, 1) else "262626")
        cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center" if j == 0 else None,
                                   indent=0 if j == 0 else 1)
        cell.fill = band
        cell.border = Border(bottom=Side(style="thin", color=LINE))
g.freeze_panes = "A8"
g.sheet_view.showGridLines = False
g.sheet_properties.tabColor = ACCENT
wb.active = 0
wb.save(OUT)
print("saved", OUT)
