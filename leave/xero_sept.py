"""Xero manual journal import files for the September AL, LSL and TIL journals,
taken line for line from the Journals tab of the September workbook."""
import csv, json, openpyxl, os
OUTD = "/home/user/Claude/leave/out"
WB = OUTD + "/2026-09_AL-LSL_Provision.xlsx"
rep = json.load(open(OUTD + "/september-build.json"))
coa = json.load(open("/root/.claude/projects/-home-user-Claude/a91f56e6-b7da-5dad-8e0b-6d5f516f9360/tool-results/mcp-Xero-get_chart_of_accounts-1791605465529.txt"))
codes = {a["code"]: a["name"] for a in coa["accounts"]}
trk = json.load(open("/root/.claude/projects/-home-user-Claude/a91f56e6-b7da-5dad-8e0b-6d5f516f9360/tool-results/mcp-Xero-get_tracking_categories-1791607765140.txt"))
jobs = {}
for c in trk["tracking_categories"]:
    if c["name"] == "Job Numbers":
        for o in c["options"]:
            jobs.setdefault(o["name"].split(" - ")[0].strip(), o["name"])
HEAD = ["*Narration", "*Date", "Description", "*AccountCode", "*TaxRate", "*Amount",
        "TrackingName1", "TrackingOption1", "TrackingName2", "TrackingOption2"]
DATE, TAX = "30/09/2026", "BAS Excluded"
ws = openpyxl.load_workbook(WB)["Journals"]

def acct(a):
    a = str(a).replace("-", "")
    assert a in codes, "account %s not in Xero" % a
    return a

def job(e):
    e = str(e).strip()
    assert e in jobs, "job %s not in Xero" % e
    return jobs[e]

def write(name, narration, lines):
    lines = [l for l in lines if round(l[2], 2) != 0]
    tot = round(sum(l[2] for l in lines), 2)
    assert tot == 0, (name, tot)
    with open(os.path.join(OUTD, name), "w", newline="") as fh:
        w = csv.writer(fh); w.writerow(HEAD)
        for desc, a, amt, e in lines:
            w.writerow([narration, DATE, desc, acct(a), TAX, "%.2f" % amt, "Job Numbers", job(e), "", ""])
    return len(lines), round(sum(l[2] for l in lines if l[2] > 0), 2)

out = {}
# ---- Annual leave: rows 5-10 expense by department, row 11 provision
al = rep["AL"]["lines"]
nar = "AL Provision - Sep 2026"
lines = []
for row, ref in zip(range(5, 11), ["BS100", "BS101", "BS102", "BS103", None, None]):
    v = al.get(ref, 0.0) if ref else 0.0
    lines.append((nar, ws["A%d" % row].value, v, ws["E%d" % row].value if not str(ws["E%d" % row].value).startswith("=") else
                  {5: "9000", 6: "9100", 7: "9105", 8: "9125", 9: "9200", 10: "9400"}[row]))
lines.append((nar, ws["A11"].value, -sum(l[2] for l in lines), ws["E11"].value))
out["AL"] = write("2026-09 AL Provision Xero Journal.csv", nar, lines)

# ---- Long service: rows 16-21 by department, row 22 provision
ls = rep["LSL"]["lines"]
nar = "LSL Provision - Sep 2026"
jobmap = {16: "9000", 17: "9100", 18: "9105", 19: "9125", 20: "9200", 21: "9400"}
lines = [(nar, ws["A%d" % r].value, ls[str(r + 14)], jobmap[r]) for r in range(16, 22)]
lines.append((nar, ws["A22"].value, -sum(l[2] for l in lines), ws["E22"].value))
out["LSL"] = write("2026-09 LSL Provision Xero Journal.csv", nar, lines)

# ---- Time in lieu: each location is three lines, salary and super debits, payable credit
loc = rep["TIL"]["by_location"]
nar = "Time in Lieu - Sep 2026"
d_by_row = {29: loc.get("PRODUCTION [PRD]", 0), 32: loc.get("VIDEO DEPT [VID]", 0), 35: loc.get("INTEGRATION [INT]", 0),
            38: 0, 41: 0, 44: loc.get("CBA Brisbane ONS [ONS]", 0), 47: loc.get("DTTL SYD FT Tech 2 [ONS]", 0),
            50: loc.get("DTTL Mel Tech FT 2 [ONS]", 0), 53: loc.get("DTTL SYD FT Tech 1 [ONS]", 0),
            56: loc.get("DTTL SYD NTSL [ONS]", 0), 59: loc.get("Bank West Contract [ONS]", 0),
            62: loc.get("OFFICE / ADMIN [CTS]", 0), 65: loc.get("DTTL Mel Tech FT 1 [ONS]", 0), 68: 0,
            71: loc.get("DTTL BNE Onsite [ONS]", 0), 74: loc.get("DTTL ADL Onsite [ONS]", 0), 77: 0,
            80: loc.get("CBA Team Leader [ONS]", 0), 83: loc.get("CBA Tech TN [ONS]", 0),
            86: loc.get("DTTL SYD Internal [ONS]", 0)}
team = {29: "PRD", 32: "VID", 35: "ONS", 38: "CONS"}
lines, adjusted = [], 0
for r, d in d_by_row.items():
    d = round(d, 2)
    if not d: continue
    sal = round(d / 1.12, 2); sup = round(sal * 0.12, 2)
    if round(sal + sup, 2) != d:
        sup = round(d - sal, 2); adjusted += 1
    desc = ("Time in Lieu - %s team as at 30/09/26" % team[r]) if r in team else nar
    e = ws["E%d" % r].value
    lines += [(desc, ws["A%d" % (r - 2)].value, sal, e), (desc, ws["A%d" % (r - 1)].value, sup, e),
              (desc, ws["A%d" % r].value, -d, e)]
out["TIL"] = write("2026-09 Time in Lieu Xero Journal.csv", nar, lines)
out["TIL_super_cent_adjustments"] = adjusted
out["TIL_total"] = round(sum(round(v, 2) for v in d_by_row.values()), 2)
print(json.dumps(out, indent=1))
