#!/usr/bin/env python3
"""Generates the PBIR report pages and visuals for the Utilisation project."""
import os, json, io, shutil, sys

HERE = os.path.dirname(os.path.abspath(__file__))
RP   = os.path.join(HERE, "pbip", "Utilisation.Report")
DEF  = os.path.join(RP, "definition")
PG   = os.path.join(DEF, "pages")

VS = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.7.0/schema.json"
PS = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.0.0/schema.json"
MS = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json"

W, H = 1920, 1080
M    = "_Measures"

def wj(p, o):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with io.open(p, "w", encoding="utf-8", newline="\n") as f:
        f.write(json.dumps(o, indent=2, ensure_ascii=False) + "\n")

def lit(v):        return {"expr": {"Literal": {"Value": v}}}
def s_(t):         return lit("'%s'" % str(t).replace("'", "''"))
def d_(n):         return lit("%sD" % n)
def b_(x):         return lit("true" if x else "false")

def meas(name, table=M):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": table}}, "Property": name}}
def col(table, name):
    return {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": name}}

def pm(name, table=M):     # projection for a measure
    return {"field": meas(name, table), "queryRef": "%s.%s" % (table, name), "nativeQueryRef": name}
def pc(table, name):       # projection for a column
    return {"field": col(table, name), "queryRef": "%s.%s" % (table, name), "nativeQueryRef": name}

def titled(text, sub=None):
    p = {"show": b_(True), "text": s_(text), "fontSize": d_(12), "bold": b_(True)}
    o = {"title": [{"properties": p}]}
    if sub:
        o["subTitle"] = [{"properties": {"show": b_(True), "text": s_(sub), "fontSize": d_(9)}}]
    return o

def V(vid, vtype, x, y, w, h, roles=None, z=0, objects=None, vco=None, sort=None, extra=None):
    v = {"$schema": VS, "name": vid,
         "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": z},
         "visual": {"visualType": vtype, "drillFilterOtherVisuals": True}}
    v["visual"]["query"] = {"queryState": roles or {}}
    if sort:    v["visual"]["query"]["sortDefinition"] = sort
    if objects: v["visual"]["objects"] = objects
    if vco:     v["visual"]["visualContainerObjects"] = vco
    if extra:   v["visual"].update(extra)
    return v

def sortby(name, direction="Descending", table=M):
    return {"sort": [{"field": meas(name, table), "direction": direction}], "isDefaultSort": True}

def slicer(vid, table, column_name, x, y, w, h, group):
    return V(vid, "slicer", x, y, w, h,
             {"Values": {"projections": [pc(table, column_name)]}},
             vco=titled(column_name),
             extra={"syncGroup": {"groupName": group, "fieldChanges": True, "filterChanges": True}})

def card(vid, measure, x, y, w=348, h=130, sub=None):
    return V(vid, "card", x, y, w, h,
             {"Values": {"projections": [pm(measure)]}},
             vco=titled(measure, sub))

# ---------------------------------------------------------------- pages
PAGES = []

# ============ 1. Summary ============
vis = []
vis.append(slicer("slicer_month",  "Dim_Date",     "Month Year",   40,  40, 300, 90, "MonthYear"))
vis.append(slicer("slicer_group",  "Dim_Employee", "Person Group", 360, 40, 300, 90, "PersonGroup"))
for i, (vid, m, sub) in enumerate([
      ("card_util",     "Utilisation %",              "Delivery staff, casuals included"),
      ("card_variance", "Utilisation vs Target (pp)", "Percentage points"),
      ("card_chg",      "Chargeable Hours (Ordinary)", None),
      ("card_avail",    "Available Hours",            "Capacity less leave"),
      ("card_cover",    "Timesheet Coverage %",       "Check this before the number above")]):
    vis.append(card(vid, m, 40 + i * 368, 150, sub=sub))
vis.append(V("chart_trend", "lineChart", 40, 300, 1120, 380,
             {"Category": {"projections": [pc("Dim_Date", "Week Ending")]},
              "Y": {"projections": [pm("Utilisation % R13W"), pm("Target Utilisation %")]}},
             vco=titled("Utilisation trend", "Rolling 13 weeks - a single fortnight is too noisy to read")))
vis.append(V("chart_group", "clusteredBarChart", 1180, 300, 700, 380,
             {"Category": {"projections": [pc("Dim_Employee", "Person Group")]},
              "Y": {"projections": [pm("Utilisation %")]}},
             sort=sortby("Utilisation %"),
             vco=titled("Utilisation by team")))
vis.append(V("chart_mix", "donutChart", 40, 700, 560, 340,
             {"Category": {"projections": [pc("Dim_PayType", "Category")]},
              "Y": {"projections": [pm("Total Hours")]}},
             vco=titled("Where the hours went")))
vis.append(V("chart_dept", "clusteredBarChart", 620, 700, 640, 340,
             {"Category": {"projections": [pc("Dim_Job", "Job Group")]},
              "Y": {"projections": [pm("Chargeable Hours (Ordinary)")]}},
             sort=sortby("Chargeable Hours (Ordinary)"),
             vco=titled("Chargeable hours by job group")))
vis.append(card("card_ot", "Overtime Hours", 1280, 700, w=290, h=160,
                sub="Outside the ratio, by design"))
vis.append(card("card_casual", "Casual Share of Delivered Hours", 1590, 700, w=290, h=160))
vis.append(card("card_dq", "Data Quality Flag", 1280, 880, w=600, h=160))
PAGES.append(("Summary", "Summary", vis))

# ============ 2. By person ============
vis = []
vis.append(slicer("p2_slicer_month", "Dim_Date", "Month Year", 40, 40, 300, 90, "MonthYear"))
vis.append(slicer("p2_slicer_scope", "Dim_Employee", "Utilisation Scope", 360, 40, 300, 90, "Scope"))
vis.append(V("tbl_people", "tableEx", 40, 150, 1840, 890,
             {"Values": {"projections": [
                 pc("Dim_Employee", "Employee Name"),
                 pc("Dim_Employee", "Person Group"),
                 pc("Dim_Employee", "Employment Type"),
                 pm("Available Hours"),
                 pm("Chargeable Hours (Ordinary)"),
                 pm("Non-Chargeable Hours"),
                 pm("Leave Hours"),
                 pm("Utilisation %"),
                 pm("Target Utilisation %"),
                 pm("Utilisation vs Target (pp)"),
                 pm("Timesheet Coverage %")]}},
             sort=sortby("Utilisation %", "Ascending"),
             vco=titled("Utilisation by person",
                        "Coverage % sits beside it on purpose - answer 'did they fill in a timesheet' first")))
PAGES.append(("By_person", "By person", vis))

# ============ 3. Where the time went ============
vis = []
vis.append(slicer("p3_slicer_month", "Dim_Date", "Month Year", 40, 40, 300, 90, "MonthYear"))
vis.append(V("p3_nonchg", "clusteredBarChart", 40, 150, 920, 440,
             {"Category": {"projections": [pc("Dim_Job", "Job Name")]},
              "Y": {"projections": [pm("Non-Chargeable Hours")]}},
             sort=sortby("Non-Chargeable Hours"),
             vco=titled("Non-chargeable hours by job", "Your list of things to kill")))
vis.append(V("p3_ot", "clusteredBarChart", 980, 150, 900, 440,
             {"Category": {"projections": [pc("Dim_Employee", "Employee Name")]},
              "Y": {"projections": [pm("Overtime Hours")]}},
             sort=sortby("Overtime Hours"),
             vco=titled("Overtime by person",
                        "Concentrated on a few names is a resourcing problem, not a utilisation one")))
vis.append(V("p3_matrix", "pivotTable", 40, 610, 1840, 430,
             {"Rows": {"projections": [pc("Dim_PayType", "Category"), pc("Dim_Job", "Job Group")]},
              "Columns": {"projections": [pc("Dim_Date", "Month Year")]},
              "Values": {"projections": [pm("Total Hours")]}},
             vco=titled("Hours by category and month")))
PAGES.append(("Where_the_time_went", "Where the time went", vis))

# ============ 4. WIP hours ============
vis = []
vis.append(slicer("p4_slicer_month", "Dim_Date", "Month Year", 40, 40, 300, 90, "MonthYear"))
vis.append(slicer("p4_slicer_basis", "Dim_Job", "Billing Basis Final", 360, 40, 300, 90, "Basis"))
for i, (vid, m, sub) in enumerate([
      ("p4_added",  "WIP Hours Added",   "This month, project jobs only"),
      ("p4_todate", "WIP Hours to Date", "Hours booked to date - NOT unbilled"),
      ("p4_unb",    "Unbilled WIP Hours","Needs Last Fully Billed Date populated"),
      ("p4_nojob",  "Hours With No Job Number", "Cannot be attributed to any job")]):
    vis.append(card(vid, m, 40 + i * 468, 150, w=448, h=150, sub=sub))
vis.append(V("p4_table", "tableEx", 40, 320, 1840, 720,
             {"Values": {"projections": [
                 pc("Dim_Job", "Job No"),
                 pc("Dim_Job", "Job Name"),
                 pc("Dim_Job", "Job Group"),
                 pc("Dim_Job", "Billing Basis Final"),
                 pm("WIP Hours Added"),
                 pm("WIP Hours to Date"),
                 pc("Dim_Job", "Last Fully Billed Date")]}},
             sort=sortby("WIP Hours Added"),
             vco=titled("WIP hours by job",
                        "Drops beside the dollar schedule on the same job key")))
PAGES.append(("WIP_hours", "WIP hours", vis))

# ============ 5. Data quality ============
vis = []
vis.append(slicer("p5_slicer_month", "Dim_Date", "Month Year", 40, 40, 300, 90, "MonthYear"))
for i, (vid, m, sub) in enumerate([
      ("p5_nojob",   "Hours With No Job Number",  "Target: 0"),
      ("p5_noemp",   "Unmatched Employee Hours",  "Target: 0"),
      ("p5_nojobm",  "Unmapped Job Hours",        "Target: 0"),
      ("p5_nopt",    "Unmapped Pay Type Hours",   "Target: 0")]):
    vis.append(card(vid, m, 40 + i * 468, 150, w=448, h=150, sub=sub))
for i, (vid, m, sub) in enumerate([
      ("p5_wdd",  "Working Days With Data",  "Compare with the next card"),
      ("p5_wdp",  "Working Days In Period",  "Unequal = partial month"),
      ("p5_age",  "Data Age (days)",         None),
      ("p5_last", "Last Timesheet Date",     "Read this before any percentage")]):
    vis.append(card(vid, m, 40 + i * 468, 320, w=448, h=150, sub=sub))
vis.append(V("p5_cover", "clusteredBarChart", 40, 490, 920, 550,
             {"Category": {"projections": [pc("Dim_Employee", "Employee Name")]},
              "Y": {"projections": [pm("Timesheet Coverage %")]}},
             sort=sortby("Timesheet Coverage %", "Ascending"),
             vco=titled("Timesheet coverage by person", "Materially below 100% means missing timesheets")))
vis.append(V("p5_unmapped", "tableEx", 980, 490, 900, 550,
             {"Values": {"projections": [
                 pc("Dim_PayType", "Pay Type"),
                 pc("Dim_PayType", "Category"),
                 pm("Total Hours")]}},
             sort=sortby("Total Hours"),
             vco=titled("Pay type mapping check", "Every row should carry a sensible category")))
PAGES.append(("Data_quality", "Data quality", vis))

# ---------------------------------------------------------------- write
if os.path.isdir(PG): shutil.rmtree(PG)
os.makedirs(PG, exist_ok=True)
order = []
for folder, display, visuals in PAGES:
    order.append(folder)
    wj(os.path.join(PG, folder, "page.json"), {
        "$schema": PS, "name": folder, "displayName": display,
        "displayOption": "FitToPage", "width": W, "height": H})
    for v in visuals:
        wj(os.path.join(PG, folder, "visuals", v["name"], "visual.json"), v)
wj(os.path.join(PG, "pages.json"), {"$schema": MS, "pageOrder": order, "activePageName": order[0]})

n = sum(len(f) for _, _, f in os.walk(PG))
print("pages: %d, visuals: %d, files: %d" % (len(PAGES), sum(len(v) for _,_,v in PAGES), n))
for folder, display, visuals in PAGES:
    print("   %-22s %-22s %d visuals" % (folder, display, len(visuals)))
