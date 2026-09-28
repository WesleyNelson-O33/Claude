"""Populates the FY27 v4 tracker from the v2 tracker and the Production Jobs list.

Order of work mirrors who would type it:
  1. Department heads - Production V (video) rows completed and split from the production
     row, Consulting job lines, Onsite invoice types / job numbers tidied.
  2. Finance - Xero Invoice No, Date and Ex GST for Onsite, Consulting and the video split.
  3. WIP Movements - the v2 WIP journal lines.

Usage: python populate_fy27.py <v4_with_user_inputs.xlsx> <v2.xlsm> <production_jobs.xlsx> <out.xlsx>
Writes a log of every change next to the output (<out>.log.txt).
"""
import datetime as dt
import re
import sys

import openpyxl
from openpyxl.utils import column_index_from_string as CI

V4, V2, JOBS, OUT = sys.argv[1:5]
log = []
wb = openpyxl.load_workbook(V4)
v2 = openpyxl.load_workbook(V2, data_only=True)["Tracker"]
jl = openpyxl.load_workbook(JOBS, data_only=True).active


def hdrmap(ws, row):
    return {c.value: c.column_letter for c in ws[row] if c.value}


def block(first_col, ncols, start=12):
    a = CI(first_col)
    rows = []
    for r in range(start, v2.max_row + 1):
        vals = [v2.cell(r, c).value for c in range(a, a + ncols)]
        if any(v not in (None, "") for v in vals[:7]):
            rows.append((r, vals))
    return rows


def d(v):
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, str):
        for f in ("%m/%d/%Y", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(v.strip(), f)
            except ValueError:
                pass
    return None


num = lambda v: v if isinstance(v, (int, float)) else None
fi = wb["Finance"]
FH = hdrmap(fi, 6)
frow = {fi[f"A{r}"].value: r for r in range(7, fi.max_row + 1)}


def fin_get(rid, h):
    return fi[f"{FH[h]}{frow[rid]}"].value


def fin_set(rid, h, v, why=""):
    cell = fi[f"{FH[h]}{frow[rid]}"]
    if cell.value != v:
        log.append(f"Finance {rid} {h}: {cell.value!r} -> {v!r} {why}")
        cell.value = v


# ============================================================= PRODUCTION
pr = wb["Production"]
PH = hdrmap(pr, 4)
# v2 production block: Date, Department, Client, Client Email, Job Number, Project Name, Event Date, Invoice Value,
# Zoho, Closed, Current Number, Posted, Xero Inv, Video Filming, Video Editing, Project Mgmt, Video PM, Prod Labour,
# Video Total, Discounts, Cross Hire, Labour, Net Total ...
v2p = [vals for _r, vals in block("CS", 27)]
# job list (workbook 1): the video split
jobs = {}
for r in range(5, jl.max_row + 1):
    job = jl[f"E{r}"].value
    if job and jl[f"A{r}"].value != "TOTAL":
        jobs[str(job)] = {h: jl.cell(r, c).value for c, h in enumerate(
            [x.value for x in jl[4]], 1)}
# main-job lookup for the package row: 26061609V -> the "26061609/..." line
by_zoho = {str(v["Zoho Number"]): v for v in jobs.values()}

prod_rows = [r for r in range(5, 5 + 1500) if pr[f"C{r}"].value not in (None, "")]
rid_of = lambda r: pr[f"{PH['Row ID']}{r}"].value
main_of_job = {}
for r in prod_rows:
    j = str(pr[f"C{r}"].value).strip()
    if not j.upper().endswith("V"):
        main_of_job.setdefault(j, []).append(r)


def v2_prod_line(job, desc):
    for vals in v2p:
        if str(vals[4]).strip() == job and str(vals[5]).strip() == str(desc).strip():
            return vals
    for vals in v2p:
        if str(vals[4]).strip() == job:
            return vals
    return None


for r in prod_rows:
    # tidy: cost centres to the list spelling, invoice type where the department left it blank
    cc = pr[f"{PH['Cost Centre']}{r}"].value
    if isinstance(cc, str) and cc.strip().upper() in ("PRODUCTION", "VIDEO") and cc != cc.strip().upper():
        pr[f"{PH['Cost Centre']}{r}"] = cc.strip().upper()
    if pr[f"{PH['Invoice Type']}{r}"].value in (None, ""):
        amt = num(pr[f"{PH['Expected Revenue Ex GST']}{r}"].value)
        pr[f"{PH['Invoice Type']}{r}"] = "Credit Note" if (amt or 0) < 0 else "Project"

video_done = []
mains_done = set()
for r in prod_rows:
    vjob = str(pr[f"C{r}"].value).strip()
    if not vjob.upper().endswith("V"):
        continue
    base = vjob[:-1]
    src = jobs.get(base) or by_zoho.get(base)
    mains = main_of_job.get(base) or [m for j, rs in main_of_job.items() if j.split("/")[0] == base for m in rs]
    if not src or not mains:
        log.append(f"Production {rid_of(r)} {vjob}: no matching job in the job list / production rows - left as typed")
        continue
    # the main row is the one whose description matches the job list project name, else the first
    main = next((m for m in mains if str(pr[f"G{m}"].value).strip() == str(src["Project Name"]).strip()), mains[0])
    video = float(src["Video Total"] or 0)
    inv_value = float(src["Invoice Value"] or 0)
    # ---- department head: V row
    pr[f"{PH['Cost Centre']}{r}"] = "VIDEO"
    pr[f"{PH['Expected Revenue Ex GST']}{r}"] = video
    pr[f"{PH['Invoice Type']}{r}"] = "Project"
    for h_v4, h_jl in [("Client Email", "Client Email"), ("Current RMS No", "Current Number"),
                       ("Job Closed", "Closed"), ("Company", "Client")]:
        if src.get(h_jl) not in (None, ""):
            pr[f"{PH[h_v4]}{r}"] = str(src[h_jl]) if h_v4 == "Current RMS No" else src[h_jl]
    if isinstance(src.get("Event Date"), dt.datetime):
        pr[f"{PH['Event Date']}{r}"] = src["Event Date"]
    # video hours move from the production row to the V row
    for h_v4, h_jl in [("Video Filming Hrs", "Video Filming"), ("Video Editing Hrs", "Video Editing"),
                       ("Video Project Mgmt Hrs", "Video Project Management")]:
        pr[f"{PH[h_v4]}{r}"] = src.get(h_jl)
        if pr[f"{PH[h_v4]}{main}"].value not in (None, ""):
            log.append(f"Production {rid_of(main)} {h_v4} {pr[f'{PH[h_v4]}{main}'].value} moved to {vjob}")
            pr[f"{PH[h_v4]}{main}"] = None
    pr[f"{PH['Notes']}{r}"] = f"Video part of {base} - Video Total per job list"
    # ---- department head: production row = invoice value less the video part
    old = pr[f"{PH['Expected Revenue Ex GST']}{main}"].value
    prod_part = round(inv_value - video, 2)
    pr[f"{PH['Expected Revenue Ex GST']}{main}"] = prod_part
    if pr[f"{PH['Cost Centre']}{main}"].value == "VIDEO":        # a Video-department job: the rest is production
        pr[f"{PH['Cost Centre']}{main}"] = "PRODUCTION"
        log.append(f"Production {rid_of(main)} {base}: cost centre VIDEO -> PRODUCTION (video now on {vjob})")
    log.append(f"Production {rid_of(main)} {base}: Expected {old} -> {prod_part} (invoice {inv_value} - video {video}); "
               f"{vjob} Expected {video}")
    # ---- finance: split the Xero invoice line between the two rows
    mrid, vrid = rid_of(main), rid_of(r)
    inv_no, inv_dt, xero = fin_get(mrid, "Xero Invoice No"), fin_get(mrid, "Xero Invoice Date"), num(fin_get(mrid, "Xero Invoiced Ex GST"))
    if xero is not None:
        fin_set(mrid, "Xero Invoiced Ex GST", round(xero - video, 2), "(video line moved to V row)")
        fin_set(vrid, "Xero Invoiced Ex GST", video)
    if inv_no not in (None, ""):
        fin_set(vrid, "Xero Invoice No", inv_no)
    if inv_dt not in (None, ""):
        fin_set(vrid, "Xero Invoice Date", inv_dt)
    video_done.append((base, inv_value, video))
    mains_done.add(main)

# the rest of the production rows: Expected = invoice value (the discount is already off the invoice)
for r in prod_rows:
    job = str(pr[f"C{r}"].value).strip()
    if job.upper().endswith("V") or r in mains_done:
        continue
    line = v2_prod_line(job, pr[f"G{r}"].value)
    if line and num(line[7]) is not None:
        old = pr[f"{PH['Expected Revenue Ex GST']}{r}"].value
        if old != line[7]:
            pr[f"{PH['Expected Revenue Ex GST']}{r}"] = line[7]
            log.append(f"Production {rid_of(r)} {job}: Expected {old} -> {line[7]} (v2 Invoice Value; discount already off)")
# dates Finance typed as 'N/A' - use the v2 date for the invoice
for r in prod_rows:
    rid = rid_of(r)
    if fin_get(rid, "Xero Invoice Date") == "N/A":
        line = v2_prod_line(str(pr[f"C{r}"].value).strip(), pr[f"G{r}"].value)
        if line and isinstance(line[0], dt.datetime):
            fin_set(rid, "Xero Invoice Date", line[0], "(was 'N/A'; v2 date)")
            fi[f"{FH['Finance Notes']}{frow[rid]}"] = "Invoice date taken from the v2 tracker - check against Xero"
# V rows copy the main row's (now fixed) invoice date
for r in prod_rows:
    job = str(pr[f"C{r}"].value).strip()
    if job.upper().endswith("V") and fin_get(rid_of(r), "Xero Invoice Date") == "N/A":
        base = job[:-1]
        m = next((m for m in mains_done if str(pr[f"C{m}"].value).split("/")[0] == base), None)
        if m:
            fin_set(rid_of(r), "Xero Invoice Date", fin_get(rid_of(m), "Xero Invoice Date"), "(follows production row)")

# APA "Package" lines: not separate Xero invoices - the value was billed in a package and comes in
# through the WIP Movements releases (as in v2), so they carry nil here to avoid counting it twice.
PACKAGE = {"25121208": "APA FYR August 2026 package - released through WIP Movements in August",
           "26061108": "APA All Hands package INV-10521 (row 26061609/...) - released through WIP Movements in August",
           "26061108V": "APA All Hands package INV-10521 (row 26061609/...) - released through WIP Movements in August"}
for r in prod_rows:
    job = str(pr[f"C{r}"].value).strip()
    rid = rid_of(r)
    if job in PACKAGE and (fin_get(rid, "Xero Invoice No") == "Package" or job.endswith("V")):
        was = pr[f"{PH['Expected Revenue Ex GST']}{r}"].value
        pr[f"{PH['Expected Revenue Ex GST']}{r}"] = 0
        pr[f"{PH['Notes']}{r}"] = f"Event value {was} billed in a package - {PACKAGE[job]}"
        fin_set(rid, "Xero Invoice No", "INV-10521" if job.startswith("26061108") else "Package", "(package)")
        fin_set(rid, "Xero Invoiced Ex GST", 0, "(package - revenue via WIP)")
        if not isinstance(fin_get(rid, "Xero Invoice Date"), dt.datetime):
            fin_set(rid, "Xero Invoice Date", dt.datetime(2026, 7, 15))
        fi[f"{FH['Finance Notes']}{frow[rid]}"] = "Not a separate invoice: " + PACKAGE[job]
        log.append(f"Production {rid} {job}: package line set to nil (was {was}); revenue comes via WIP Movements")

# ================================================================= ONSITE
on = wb["Onsite"]
OH = hdrmap(on, 4)
v2s = [vals for _r, vals in block("BQ", 7)]   # Date, Client, Invoice Number, Job Number, Description, Amount, Notes
on_rows = [r for r in range(5, 1505) if any(on[f"{OH[h]}{r}"].value not in (None, "") for h in ("Client", "Job Description", "Job Number"))]
used = set()
JOB_FIX = {"(no deductions for Sept)": "2846/1401/725/1400/1530/1862/2402/2849/974",
           "(No half-tech role for Sept 2026)": "732/1144",
           "2604901 (PWC Brisbane AV Support - 12 Months)": "2604901"}
for r in on_rows:
    job, desc, client = on[f"C{r}"].value, on[f"G{r}"].value, on[f"F{r}"].value
    exp = on[f"{OH['Expected Revenue Ex GST']}{r}"].value
    match = None
    for i, vals in enumerate(v2s):
        if i in used:
            continue
        if str(vals[4] or "").strip() == str(desc or "").strip() and str(vals[1] or "").strip() == str(client or "").strip() \
                and (vals[5] == exp or exp in (None, "") or vals[5] in (None, "")):
            match = i
            break
    if match is None:
        log.append(f"Onsite {on[f'X{r}'].value} ({client} / {desc}): no v2 line found - Finance left blank")
        continue
    used.add(match)
    vals = v2s[match]
    rid = on[f"{OH['Row ID']}{r}"].value
    # department head tidy-ups
    if job in JOB_FIX:
        on[f"C{r}"] = JOB_FIX[job]
        note = on[f"{OH['Notes']}{r}"].value
        on[f"{OH['Notes']}{r}"] = f"{job}" + (f" | {note}" if note else "")
        log.append(f"Onsite {rid}: Job Number {job!r} -> {JOB_FIX[job]!r} (text moved to Notes)")
    if vals[6] and not on[f"{OH['Notes']}{r}"].value:
        on[f"{OH['Notes']}{r}"] = str(vals[6])[:250]
    if on[f"{OH['Invoice Type']}{r}"].value in (None, ""):
        text = f"{desc or ''}".upper()
        contract = bool(re.search(r"CONTRACT|JULY 2026|AUGUST 2026|SEPTEMBER 2026|CSJUN27|BWPCS", text)) \
            and "ADDITIONAL" not in text
        on[f"{OH['Invoice Type']}{r}"] = "Contract" if contract else "Ad-hoc"
    # finance
    if vals[2]:
        fin_set(rid, "Xero Invoice No", str(vals[2]).strip())
    if isinstance(vals[0], dt.datetime):
        fin_set(rid, "Xero Invoice Date", vals[0])
    if num(vals[5]) is not None and vals[2]:
        fin_set(rid, "Xero Invoiced Ex GST", vals[5])
unmatched_v2 = [v for i, v in enumerate(v2s) if i not in used]
for v in unmatched_v2:
    log.append(f"v2 Onsite line not on the v4 Onsite sheet: {v[:6]}")

# ============================================================= CONSULTING
co = wb["Consulting"]
CH = hdrmap(co, 4)
v2c = [vals for _r, vals in block("DU", 18)]
# Date, Status, Department, Client, Job, Project, Link, Notes, Inv No, Inv Value, Lab Rev, Eq Rev, Sub Rev,
# Lab Exp, Eq Exp, Sub Exp, Margin, Margin %
r = 5
for vals in v2c:
    (date, status, dept, client, job, proj, link, notes, inv, value, lr, er, sr, lx, ex, sx) = vals[:16]
    rid = co[f"{CH['Row ID']}{r}"].value
    text = f"{proj} {notes}".upper()
    itype = ("Progress Claim" if "PROGRESS CLAIM" in text else
             "Subscription" if re.search(r"LICEN|SUBSCRIPTION", text) else "Project")
    put = {"Job Number": str(job).strip(), "Client": str(client).strip(), "Job Description": str(proj).strip(),
           "Cost Centre": str(dept).strip().upper(), "Invoice Type": itype, "Tax Code": "GST 10%",
           "Expected Revenue Ex GST": num(value), "Qwilr Quote": link,
           "Labour Revenue": num(lr), "Equipment Revenue": num(er), "Subscription Revenue": num(sr),
           "Labour Expense (External)": num(lx), "Equipment Expense (Internal)": num(ex),
           "Subscription & Licences Expense": num(sx),
           "Notes": " | ".join(x for x in [f"Status: {status}" if status else "", str(notes or "").strip()] if x)}
    for h, v in put.items():
        if v not in (None, ""):
            co[f"{CH[h]}{r}"] = v
    if inv:
        fin_set(rid, "Xero Invoice No", str(inv).strip())
        if isinstance(date, dt.datetime):
            fin_set(rid, "Xero Invoice Date", date)
        if num(value) is not None:
            fin_set(rid, "Xero Invoiced Ex GST", value)
    r += 1
log.append(f"Consulting: {len(v2c)} job lines entered from v2")

# ================================================================== WIP
wp = wb["WIP Movements"]
WH = hdrmap(wp, 4)
v2w = [vals for _r, vals in block("EW", 7)]    # Date, (blank), Department, Client, Job, Description, WIP Total
CCMAP = {"PRODUCTION": "PRODUCTION", "PRD": "PRODUCTION", "VIDEO": "VIDEO", "VID": "VIDEO", "INTEGRATION": "INTEGRATION",
         "INT": "INTEGRATION", "ONSITE": "ONSITE", "CONSULTING": "CONSULTING"}
r = 5
skipped = 0
for vals in v2w:
    date, _x, dept, client, job, desc, amt = vals
    # the v2 WIP block stores Job Number before Department: Date | Job | Department | Client | Description | Total
    date, job, dept, client, desc, amt = vals[0], vals[1] if vals[1] else vals[4], vals[2], vals[3], vals[4], vals[5]
    if num(amt) is None or date is None:
        skipped += 1
        log.append(f"WIP line skipped (no date or amount): {vals}")
        continue
    month = dt.datetime(date.year, date.month, 1) + dt.timedelta(days=32)
    month = dt.datetime(month.year, month.month, 1) - dt.timedelta(days=1)
    dtext = str(desc).upper()
    if "PRE-PROD" in dtext or "PRE WORK" in dtext or "PRE-PRD" in dtext or "PRE-VID" in dtext or "PRE-PRODUCTION" in dtext or "INDICATIVE" in dtext:
        wtype = "Accrual - unbilled work" if amt > 0 else "Reversal of prior accrual"
    elif "NOT INVOICED" in dtext:
        wtype = "Accrual - unbilled work"
    else:
        wtype = "Release of deferral" if amt > 0 else "Deferral - invoiced in advance"
    row = {"Month": month, "Job Number": str(job).strip(), "Cost Centre": CCMAP.get(str(dept).strip().upper(), None),
           "Description": str(desc).strip(), "Type": wtype, "Revenue or Cost": "Revenue", "Amount": amt}
    for h, v in row.items():
        if v not in (None, ""):
            wp[f"{WH[h]}{r}"] = v
    r += 1
log.append(f"WIP Movements: {r - 5} lines entered from v2, {skipped} skipped")

wb.save(OUT)
open(OUT + ".log.txt", "w").write("\n".join(log))
print("saved", OUT, "| log lines", len(log), "| video rows", len(video_done), "| onsite unmatched v2", len(unmatched_v2))
