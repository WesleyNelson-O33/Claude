"""September 2026 WIP journals for Xero import, from the FY27 tracker's WIP Movements tab.

Same layout as August's WIP journal in the recording: one pair per item, the
P&L line then the 11300 line, same description, cost centre and job on both,
BAS Excluded, dated the last day of the month.
"""
import csv, json, sys, datetime, openpyxl

TRACKER, TRACKING, OUTDIR = sys.argv[1], sys.argv[2], sys.argv[3]
jobs = [o["name"] for c in json.load(open(TRACKING))["tracking_categories"] if c["name"] == "Job Numbers" for o in c["options"]]
def job_option(j):
    hit = [o for o in jobs if o.startswith(str(j).strip() + " - ")]
    assert len(hit) == 1, (j, hit); return hit[0]

ws = openpyxl.load_workbook(TRACKER, read_only=True, data_only=False)["WIP Movements"]
hdr = None; items = []
for r, row in enumerate(ws.iter_rows(values_only=True), 1):
    if r == 4: hdr = row; continue
    if r < 5: continue
    d = dict(zip(hdr, row))
    m = d["Month"]
    if isinstance(m, (int, float)): m = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(m))
    if str(m)[:7] == "2026-09" and d["Amount"]:
        d["row"] = r; items.append(d)

# Bankwest INV-10645 is three revenue accounts on the invoice
SPLIT = {"26092101": [("42100", 400), ("42150", 400), ("42500", 150)]}
NARR, DATE = "WIP Journal September 2026", "30/09/2026"
main, auto, skipped = [], [], []
for d in items:
    notes = str(d["Notes"] or "")
    if d["Type"] == "Reversal of prior accrual" and "Auto-reverse" in notes:
        skipped.append(d); continue          # Xero already reversed August's auto-reversing journal
    dest = auto if "Auto-reverse" in notes else main
    amt = float(d["Amount"]); cc = d["Cost Centre"]; job = job_option(d["Job Number"]); desc = d["Description"].strip()
    parts = SPLIT.get(str(d["Job Number"]).strip(), [(str(d["P&L Account"]), abs(amt))])
    sign = 1 if amt > 0 else -1
    for acct, a in parts:
        dest.append([NARR, DATE, desc, acct, "BAS Excluded", round(-sign * a, 2), "Cost Centres", cc, "Job Numbers", job])
    dest.append([NARR, DATE, desc, "11300", "BAS Excluded", round(sign * abs(amt), 2), "Cost Centres", cc, "Job Numbers", job])

HEAD = ["*Narration", "*Date", "Description", "*AccountCode", "*TaxRate", "*Amount",
        "TrackingName1", "TrackingOption1", "TrackingName2", "TrackingOption2"]
for name, lines in (("2026-09 WIP Journal Xero.csv", main), ("2026-09 WIP Journal Xero - AUTO-REVERSE.csv", auto)):
    assert abs(sum(l[5] for l in lines)) < 0.005, name
    with open(f"{OUTDIR}/{name}", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(HEAD); w.writerows(lines)
    print(name, len(lines), "lines, debits", round(sum(l[5] for l in lines if l[5] > 0), 2))
    for l in lines: print("   ", l[3], l[5], l[7], l[9][:30], l[2][:50])
print("not posted (already auto-reversed in Xero):", [(d["row"], d["Job Number"], d["Amount"]) for d in skipped])
