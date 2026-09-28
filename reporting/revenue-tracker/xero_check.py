"""Checks the Finance sheet against a Xero invoice export (JSON from the Xero connector).

Usage: python xero_check.py <tracker_recalculated.xlsx> <xero_invoices.json> <out.xlsx> <first_inv_no> <from> <to>
Only invoices numbered from first_inv_no, and dated from..to, are compared, so older-year invoices
on the tracker are not reported as missing.
"""
import json
import re
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

TR, XJ, OUT, FIRST, D0, D1 = sys.argv[1:7]
FIRST = int(FIRST)


def num(s):
    m = re.search(r"(\d{4,6})", str(s))
    return int(m.group(1)) if m else None


X = json.load(open(XJ))
xero = {}          # exact Xero invoice number -> record
by_num = defaultdict(list)
for i in X:
    jobs = sorted(set(re.findall(r"#\s*:?\s*#?(\d{6,10}V?)", " ".join((l.get("description") or "") for l in i.get("line_items", [])) + " " + (i.get("reference") or ""))))
    rec = dict(no=i["invoice_number"], date=i["invoice_date"], net=float(i["amount_net"]),
               cred_inc=float(i.get("amount_credited") or 0), client=i["contact"]["name"],
               ref=(i.get("reference") or "").strip(), jobs=", ".join(jobs), status=i["status"])
    xero[i["invoice_number"]] = rec
    by_num[num(i["invoice_number"])].append(rec)

wb = openpyxl.load_workbook(TR, data_only=True)
ws = wb["Finance"]
h = {c.value: c.column for c in ws[6] if c.value}
rows = defaultdict(list)       # invoice number (int) -> tracker rows typed with that one number
multi = defaultdict(list)      # cell text -> tracker rows that typed several invoice numbers in one cell
credit_rows = []
for r in range(7, ws.max_row + 1):
    raw = ws.cell(r, h["Xero Invoice No"]).value
    if not raw:
        continue
    d = ws.cell(r, h["Xero Invoice Date"]).value
    row = dict(id=ws.cell(r, 1).value, dept=ws.cell(r, 2).value, job=ws.cell(r, 3).value, client=ws.cell(r, 4).value,
               desc=ws.cell(r, 5).value, raw=str(raw), date=d.date().isoformat() if d else "",
               amt=ws.cell(r, h["Xero Invoiced Ex GST"]).value or 0)
    if str(raw).strip().upper().startswith("CN"):
        credit_rows.append(row)            # credit notes are not in the Xero invoice pull
        continue
    nums = [int(n) for n in re.findall(r"\d{4,6}", str(raw))]
    if len(nums) > 1:
        multi[str(raw)].append(row)
    elif nums:
        rows[nums[0]].append(row)

issues = []   # (priority, type, invoice, xero date, xero ex gst, tracker ex gst, difference, rows, what to do)


def add(p, t, inv, xdate, xnet, ta, rs, todo):
    issues.append((p, t, inv, xdate, xnet, ta, round((ta or 0) - (xnet or 0), 2), ", ".join(r["id"] for r in rs), todo))


def pick(xs, rs):
    """Where Xero has two invoices with one number, take the one whose customer matches the tracker client."""
    if len(xs) == 1 or not rs:
        return xs[0] if xs else None
    for x in xs:
        if any(str(r["client"] or "").split()[0].lower() in x["client"].lower() for r in rs if r["client"]):
            return x
    return xs[0]


covered = set()
for raw, rs in multi.items():
    nums = [int(n) for n in re.findall(r"\d{4,6}", raw)]
    xs = [pick(by_num[n], rs) for n in nums if by_num.get(n)]
    covered.update(x["no"] for x in xs)
    ta, xn = round(sum(r["amt"] for r in rs), 2), round(sum(x["net"] for x in xs), 2)
    add(3 if abs(ta - xn) < 0.005 else 2, "Two invoice numbers typed in one cell", " + ".join(x["no"] for x in xs),
        ", ".join(sorted({x["date"] for x in xs})), xn, ta, rs,
        f"Cell reads {raw!r}. Totals {'agree' if abs(ta - xn) < 0.005 else 'DO NOT agree'}. Put each invoice on its own row so each ties to Xero")

for n in sorted({k for k in by_num if k and k >= FIRST} | {k for k in rows if k >= FIRST}):
    xs, rs = by_num.get(n, []), rows.get(n, [])
    if len(xs) > 1:
        add(1, "Xero has two invoices with this number", " / ".join(repr(x["no"]) for x in xs), "", None, None, [],
            "; ".join(f'{x["no"]!r} {x["client"]} ${x["net"]:,.2f} {x["date"]}' for x in xs) +
            " - renumber one in Xero so every invoice number is unique")
    if not xs:
        add(1, "On the tracker as invoiced, but not an approved invoice in Xero", f"INV-{n}", "", None,
            round(sum(r["amt"] for r in rs), 2), rs,
            "Draft, voided or wrong number. Leave the invoice no and amount blank until it is approved in Xero, or fix the number")
        continue
    if not rs:
        for x in xs:
            if x["no"] not in covered:
                add(1, "In Xero, missing from the tracker", x["no"], x["date"], x["net"], None, [],
                    f'{x["client"]} | {x["ref"]} | job {x["jobs"] or "?"} - department adds the job row, finance types the invoice')
        continue
    x = pick(xs, rs)
    ta = round(sum(r["amt"] for r in rs), 2)
    if abs(ta - x["net"]) > 0.005:
        why = f'Xero: {x["client"]} | {x["ref"]} | job {x["jobs"] or "not on the invoice"}'
        if x["cred_inc"]:
            why += f' | Xero shows ${x["cred_inc"]:,.2f} inc GST already credited against this invoice'
        add(2, "Amount differs", x["no"], x["date"], x["net"], ta, rs, why)
    ds = {r["date"] for r in rs if r["date"]}
    if any(not r["date"] for r in rs):
        add(2 if x["date"][:7] else 3, "No invoice date on the tracker", x["no"], x["date"], x["net"], ta,
            [r for r in rs if not r["date"]], f'Type {x["date"]}')
    if ds and ds != {x["date"]}:
        moved = any(d[:7] != x["date"][:7] for d in ds)
        add(2 if moved else 3, "Date differs - MOVES MONTH" if moved else "Date differs (same month)", x["no"], x["date"],
            x["net"], ta, rs, f'Tracker {", ".join(sorted(ds))} -> Xero {x["date"]}')
    bad = [r for r in rs if r["raw"].strip() != x["no"]]
    if bad:
        add(3, "Invoice number typed differently from Xero", x["no"], x["date"], x["net"], ta, bad,
            f'Tracker {bad[0]["raw"]!r} -> type {x["no"]!r}')
for r in credit_rows:
    add(3, "Credit note row - check against Xero credit notes (not in this pull)", r["raw"], "", None, r["amt"], [r],
        "Credit notes are not in the invoice pull. If the credit is already netted off the invoice amount above, this row double counts it")

issues.sort(key=lambda t: (t[0], t[1], t[2]))
out = openpyxl.Workbook()
o = out.active
o.title = "Xero check"
o["A1"] = f"Finance sheet vs Xero - invoices INV-{FIRST} onwards dated {D0} to {D1}"
o["A1"].font = Font(name="Arial", size=14, bold=True, color="1F3864")
o["A2"] = ("Source: Xero connector, approved (AUTHORISED/PAID) sales invoices only. Drafts are not visible to the check, "
           "so a 'not in Xero' line may be a draft that has not been approved yet.")
o["A2"].font = Font(name="Arial", size=9, italic=True)
heads = ["Priority", "Issue", "Invoice", "Xero Date", "Xero Ex GST", "Tracker Ex GST", "Tracker less Xero", "Tracker Rows", "What to do"]
for c, t in enumerate(heads, 1):
    cell = o.cell(4, c, t)
    cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F3864")
    cell.alignment = Alignment(wrap_text=True, vertical="center")
for i, t in enumerate(issues):
    for c, v in enumerate(t, 1):
        cell = o.cell(5 + i, c, v)
        cell.font = Font(name="Arial", size=10, bold=c == 1 and v == 1)
        if c in (5, 6, 7):
            cell.number_format = '$#,##0.00;[Red]-$#,##0.00;"-"'
        cell.alignment = Alignment(wrap_text=c in (2, 9), vertical="top")
    if t[0] == 1:
        for c in range(1, 10):
            o.cell(5 + i, c).fill = PatternFill("solid", fgColor="FCE4E4")
for L, w in zip("ABCDEFGHI", (8, 34, 24, 11, 13, 13, 13, 22, 90)):
    o.column_dimensions[L].width = w
o.freeze_panes = "A5"
o.auto_filter.ref = f"A4:I{4 + len(issues)}"
out.save(OUT)
for t in issues:
    print(t[0], "|", t[1], "|", t[2], "|", t[4], "|", t[5], "|", t[6], "|", t[7], "|", t[8][:110])
