#!/usr/bin/env python3
"""
Builds Utilisation.pbit — a Power BI Template carrying the full semantic model.

A .pbix cannot be hand-authored: its data model is a compiled Analysis Services
partition. A .pbit can, because its model is plain TMSL JSON. Power BI Desktop
opens the template, builds the model, and you save it as .pbix from there.

Run:  python3 build_pbit.py [output_dir]
"""
import json, zipfile, sys, os, uuid

OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
PBIT = os.path.join(OUT, "Utilisation.pbit")

# ---------------------------------------------------------------- M queries
M = {}

M["Dim_Date"] = r'''
let
    Start    = #date(FY_Start_Year, 7, 1),
    End      = #date(FY_End_Year + 1, 6, 30),
    Dates    = List.Dates(Start, Duration.Days(End - Start) + 1, #duration(1,0,0,0)),
    T        = Table.FromList(Dates, Splitter.SplitByNothing(), {"Date"}),
    Typed    = Table.TransformColumnTypes(T, {{"Date", type date}}),
    A1  = Table.AddColumn(Typed,"Date Key",     each Date.Year([Date])*10000 + Date.Month([Date])*100 + Date.Day([Date]), Int64.Type),
    A2  = Table.AddColumn(A1,   "Year",         each Date.Year([Date]), Int64.Type),
    A3  = Table.AddColumn(A2,   "Month No",     each Date.Month([Date]), Int64.Type),
    A4  = Table.AddColumn(A3,   "Month",        each Date.ToText([Date],[Format="MMM",Culture="en-AU"]), type text),
    A5  = Table.AddColumn(A4,   "Month Sort",   each Date.Year([Date])*100 + Date.Month([Date]), Int64.Type),
    A6  = Table.AddColumn(A5,   "Month Year",   each Date.ToText([Date],[Format="MMM yyyy",Culture="en-AU"]), type text),
    A7  = Table.AddColumn(A6,   "FY",           each if Date.Month([Date]) >= 7 then Date.Year([Date])+1 else Date.Year([Date]), Int64.Type),
    A8  = Table.AddColumn(A7,   "FY Label",     each "FY" & Text.From([FY]), type text),
    A9  = Table.AddColumn(A8,   "FY Month No",  each if Date.Month([Date]) >= 7 then Date.Month([Date])-6 else Date.Month([Date])+6, Int64.Type),
    A10 = Table.AddColumn(A9,   "FY Quarter",   each "Q" & Text.From(Number.RoundUp([FY Month No]/3)), type text),
    A11 = Table.AddColumn(A10,  "Week Ending",  each Date.EndOfWeek([Date], Day.Monday), type date),
    A12 = Table.AddColumn(A11,  "Day Name",     each Date.ToText([Date],[Format="ddd",Culture="en-AU"]), type text),
    A13 = Table.AddColumn(A12,  "Day of Week No", each Date.DayOfWeek([Date], Day.Monday)+1, Int64.Type),
    A14 = Table.AddColumn(A13,  "Is Weekday",   each Date.DayOfWeek([Date], Day.Monday) < 5, type logical)
in
    A14
'''

M["Dim_Employee"] = r'''
let
    Src      = Csv.Document(File.Contents(Text.TrimEnd(MappingFolderPath,"\") & "\map_employees.csv"),
                            [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Head     = Table.PromoteHeaders(Src, [PromoteAllScalars=true]),
    Typed    = Table.TransformColumnTypes(Head, {
                  {"Employee Key", type text}, {"Employee Id", type text}, {"Employee Name", type text},
                  {"Person Group", type text}, {"Employment Type", type text},
                  {"First Seen", type date}, {"Last Seen", type date},
                  {"Utilisation Scope", type text}, {"Standard Weekly Hours", type number},
                  {"FTE", type number}, {"Target Utilisation %", type number}}, "en-AU"),
    // sensible defaults so the model works before every cell is filled in
    D1 = Table.ReplaceValue(Typed, null, "Delivery", Replacer.ReplaceValue, {"Utilisation Scope"}),
    D2 = Table.ReplaceValue(D1,    null, 38,         Replacer.ReplaceValue, {"Standard Weekly Hours"}),
    D3 = Table.ReplaceValue(D2,    null, 1,          Replacer.ReplaceValue, {"FTE"}),
    D4 = Table.ReplaceValue(D3,    null, 0,          Replacer.ReplaceValue, {"Target Utilisation %"}),
    IsCasual = Table.AddColumn(D4, "Is Casual", each Text.Lower([Employment Type]) = "casual", type logical),
    Keep     = Table.SelectRows(IsCasual, each [Employee Key] <> null and [Employee Key] <> ""),
    Unique   = Table.Distinct(Keep, {"Employee Key"})
in
    Unique
'''

M["Dim_Job"] = r'''
let
    Src   = Csv.Document(File.Contents(Text.TrimEnd(MappingFolderPath,"\") & "\map_jobs.csv"),
                         [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Head  = Table.PromoteHeaders(Src, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Head, {
               {"Job No", Int64.Type}, {"Job Name", type text}, {"Chargeable", type text},
               {"Billing Basis", type text}, {"Job Group", type text},
               {"First Booked", type date}, {"Last Booked", type date},
               {"Last Fully Billed Date", type date}, {"Override Billing Basis", type text}}, "en-AU"),
    // a manual override always wins over the derived basis
    Final = Table.AddColumn(Typed, "Billing Basis Final",
               each if [Override Billing Basis] <> null and Text.Trim([Override Billing Basis]) <> ""
                    then Text.Trim([Override Billing Basis]) else [Billing Basis], type text),
    Uniq  = Table.Distinct(Final, {"Job No"})
in
    Uniq
'''

M["Dim_PayType"] = r'''
let
    Src   = Csv.Document(File.Contents(Text.TrimEnd(MappingFolderPath,"\") & "\map_paytypes.csv"),
                         [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Head  = Table.PromoteHeaders(Src, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Head, {
               {"Pay Type", type text}, {"Pay Type Trimmed", type text},
               {"Category", type text}, {"Counts Toward Capacity", type text}}, "en-AU"),
    Key   = Table.AddColumn(Typed, "Pay Type Key",
               each Text.Upper(Text.Combine(List.Select(Text.Split(Text.Trim([Pay Type])," "), each _ <> ""), " ")), type text),
    Flag  = Table.AddColumn(Key, "Counts Toward Capacity Flag",
               each Text.Lower(Text.Trim([Counts Toward Capacity])) = "yes", type logical),
    Uniq  = Table.Distinct(Flag, {"Pay Type Key"})
in
    Uniq
'''

M["Fact_Timesheet"] = r'''
let
    Book    = Excel.Workbook(File.Contents(UtilWorkbookPath), null, true),
    Sheet   = Book{[Item="Raw data", Kind="Sheet"]}[Data],
    // the header sits on row 2 of that sheet, not row 1
    Skip    = Table.Skip(Sheet, 1),
    Head    = Table.PromoteHeaders(Skip, [PromoteAllScalars=true]),
    Ren     = Table.RenameColumns(Head, {
                 {" Company Client External Id", "Job No"},
                 {"Company Client",              "Job Name Raw"},
                 {"Counter Source Date",         "Work Date"},
                 {"Employee Id",                 "Employee Id"},
                 {"Worked Hours",                "Hours"},
                 {"Name",                        "Employee Name"},
                 {"Chargeable Category",         "Chargeable Category"},
                 {"Pay Type",                    "Pay Type"}}, MissingField.Ignore),
    Typed   = Table.TransformColumnTypes(Ren, {
                 {"Work Date", type date}, {"Hours", type number},
                 {"Employee Id", type text}, {"Employee Name", type text},
                 {"Chargeable Category", type text}, {"Pay Type", type text}}, "en-AU"),
    HasDate = Table.SelectRows(Typed, each [Work Date] <> null and [Hours] <> null and [Hours] <> 0),
    JobNum  = Table.AddColumn(HasDate, "Job No Int",
                 each try Number.From(Text.Trim(Text.From([Job No]))) otherwise null, Int64.Type),
    DateKey = Table.AddColumn(JobNum, "Date Key",
                 each Date.Year([Work Date])*10000 + Date.Month([Work Date])*100 + Date.Day([Work Date]), Int64.Type),
    PtKey   = Table.AddColumn(DateKey, "Pay Type Key",
                 each if [Pay Type] = null then ""
                      else Text.Upper(Text.Combine(List.Select(Text.Split(Text.Trim([Pay Type])," "), each _ <> ""), " ")), type text),
    // Employee Key resolves ID collisions. map_employees.csv holds one row per
    // (Employee Id, Employee Name) pair, so joining on BOTH columns is an exact
    // match - and it is a single hash join rather than a per-row scan, which
    // matters at 47,000 rows.
    IdNorm  = Table.AddColumn(PtKey, "Emp Id Norm",
                 each Text.Upper(Text.Trim(Text.From([Employee Id]))), type text),
    NameNorm= Table.AddColumn(IdNorm, "Emp Name Norm",
                 each Text.Trim(Text.From([Employee Name])), type text),
    EmpSrc  = Table.Distinct(
                 Table.SelectColumns(
                   Table.AddColumn(
                     Table.AddColumn(Dim_Employee, "Emp Id Norm",
                       each Text.Upper(Text.Trim(Text.From([Employee Id]))), type text),
                     "Emp Name Norm", each Text.Trim(Text.From([Employee Name])), type text),
                   {"Emp Id Norm","Emp Name Norm","Employee Key"}),
                 {"Emp Id Norm","Emp Name Norm"}),
    Joined  = Table.NestedJoin(NameNorm, {"Emp Id Norm","Emp Name Norm"},
                               EmpSrc,   {"Emp Id Norm","Emp Name Norm"},
                               "EmpMatch", JoinKind.LeftOuter),
    EmpKey  = Table.ExpandTableColumn(Joined, "EmpMatch", {"Employee Key"}, {"Employee Key"}),
    Chg     = Table.AddColumn(EmpKey, "Chargeable Hours",
                 each if Text.Lower(Text.Trim(if [Chargeable Category] = null then "" else [Chargeable Category])) = "chargeable" then [Hours] else 0, type number),
    Sel     = Table.SelectColumns(Chg, {
                 "Employee Key","Employee Id","Employee Name","Date Key","Work Date",
                 "Job No Int","Job Name Raw","Pay Type","Pay Type Key",
                 "Chargeable Category","Hours","Chargeable Hours"}),
    Fin     = Table.RenameColumns(Sel, {{"Job No Int","Job No"}})
in
    Fin
'''

M["Fact_Capacity"] = r'''
let
    Emp   = Table.SelectRows(
              Table.SelectColumns(Dim_Employee,
                {"Employee Key","Standard Weekly Hours","FTE","First Seen","Last Seen","Is Casual"}),
              each [Is Casual] = false),
    Cal   = Table.SelectColumns(Table.SelectRows(Dim_Date, each [Is Weekday] = true), {"Date","Date Key"}),
    Cross = Table.AddColumn(Emp, "Cal", each Cal),
    Exp   = Table.ExpandTableColumn(Cross, "Cal", {"Date","Date Key"}),
    // only days the person was actually with the business
    Serv  = Table.SelectRows(Exp, each [Date] >= [First Seen] and [Date] <= [Last Seen]),
    Cap   = Table.AddColumn(Serv, "Capacity Hours",
              each ((if [Standard Weekly Hours] = null then 0 else [Standard Weekly Hours]) / 5)
                   * (if [FTE] = null then 1 else [FTE]), type number),
    Fin   = Table.SelectColumns(Cap, {"Employee Key","Date Key","Date","Capacity Hours"})
in
    Fin
'''

# ---------------------------------------------------------------- measures
# (name, expression, format string, description)
MEASURES = [
 ("Total Hours", "SUM ( Fact_Timesheet[Hours] )", "#,0.0", "Every timesheet hour, including leave and public holidays."),
 ("Chargeable Hours", "SUM ( Fact_Timesheet[Chargeable Hours] )", "#,0.0", "Hours on jobs classified Chargeable."),
 ("Chargeable Hours (Ordinary)",
  "CALCULATE ( [Chargeable Hours], Dim_PayType[Category] = \"Ordinary\" )", "#,0.0",
  "Chargeable hours at ordinary time. This is the utilisation numerator - overtime sits outside the ratio."),
 ("Non-Chargeable Hours",
  "CALCULATE ( [Total Hours] - [Chargeable Hours], Dim_PayType[Category] = \"Ordinary\" )", "#,0.0",
  "Ordinary worked hours that were not chargeable. Excludes leave, public holidays and overtime."),
 ("Leave Hours", "CALCULATE ( [Total Hours], Dim_PayType[Category] = \"Leave\" )", "#,0.0", ""),
 ("Public Holiday Hours", "CALCULATE ( [Total Hours], Dim_PayType[Category] = \"Public Holiday\" )", "#,0.0", ""),
 ("Overtime Hours", "CALCULATE ( [Total Hours], Dim_PayType[Category] = \"Overtime\" )", "#,0.0",
  "Reported beside the ratio, never inside it. Overtime is excluded from both the numerator and the denominator."),
 ("Worked Hours", "[Chargeable Hours (Ordinary)] + [Non-Chargeable Hours]", "#,0.0", "Actual work. Excludes leave, public holidays and overtime."),

 # ---- capacity ----
 ("Permanent Capacity Hours", "SUM ( Fact_Capacity[Capacity Hours] )", "#,0.0",
  "Contracted hours for permanents, on working days they were employed."),
 ("Casual Capacity Hours",
  "CALCULATE ( [Worked Hours], Dim_Employee[Is Casual] = TRUE () )", "#,0.0",
  "For a casual, capacity is the hours you engaged them for. You carry no idle cost for a casual you did not call in."),
 ("Capacity Hours",
  "[Permanent Capacity Hours] + [Casual Capacity Hours]", "#,0.0", ""),
 ("Available Hours",
  "VAR Cap = [Capacity Hours]\nVAR Lv = [Leave Hours]\nRETURN MAX ( Cap - Lv, 0 )", "#,0.0",
  "Capacity less leave taken. Public holidays are NOT subtracted again - they are already excluded from Fact_Capacity."),

 # ---- the ratios ----
 ("Utilisation %",
  "DIVIDE (\n    CALCULATE ( [Chargeable Hours (Ordinary)], Dim_Employee[Utilisation Scope] = \"Delivery\" ),\n"
  "    CALCULATE ( [Available Hours],   Dim_Employee[Utilisation Scope] = \"Delivery\" )\n)", "0.0%",
  "HEADLINE. Delivery staff only, casuals included. Overhead staff are excluded from the ratio but their hours are still reported."),
 ("Utilisation % (All Staff)",
  "DIVIDE ( [Chargeable Hours (Ordinary)], [Available Hours] )", "0.0%",
  "Every person including overhead roles. Materially lower - use for capacity planning, not performance."),
 ("Billable Mix %", "DIVIDE ( [Chargeable Hours (Ordinary)], [Worked Hours] )", "0.0%",
  "Of the work actually done, how much was chargeable. Immune to timesheet-completion gaps."),
 ("Productive %", "DIVIDE ( [Worked Hours], [Available Hours] )", "0.0%", ""),

 # ---- targets ----
 ("Target Utilisation %",
  "DIVIDE (\n    SUMX ( Dim_Employee, [Available Hours] * Dim_Employee[Target Utilisation %] ),\n"
  "    SUMX ( Dim_Employee, [Available Hours] )\n)", "0.0%",
  "Availability-weighted, so a part-timer moves the group target in proportion to their capacity."),
 ("Target Chargeable Hours",
  "SUMX ( Dim_Employee, [Available Hours] * Dim_Employee[Target Utilisation %] )", "#,0.0", ""),
 ("Utilisation vs Target (pp)", "[Utilisation %] - [Target Utilisation %]", "+0.0;-0.0;0.0",
  "Percentage POINTS above or below target, not a percentage change."),

 # ---- casual signal ----
 ("Casual Share of Delivered Hours",
  "DIVIDE ( CALCULATE ( [Worked Hours], Dim_Employee[Is Casual] = TRUE () ), [Worked Hours] )", "0.0%",
  "How much of the work is carried by casual labour."),
 ("Casual Non-Chargeable %",
  "CALCULATE ( DIVIDE ( [Non-Chargeable Hours], [Worked Hours] ), Dim_Employee[Is Casual] = TRUE () )", "0.0%",
  "Paying casuals to do internal work is the expensive failure mode. Casual utilisation is ~100% by construction and tells you nothing."),

 # ---- time intelligence ----
 ("Utilisation % R13W",
  "VAR W = DATESINPERIOD ( Dim_Date[Date], MAX ( Dim_Date[Date] ), -91, DAY )\n"
  "RETURN DIVIDE (\n"
  "    CALCULATE ( [Chargeable Hours (Ordinary)], W, Dim_Employee[Utilisation Scope] = \"Delivery\" ),\n"
  "    CALCULATE ( [Available Hours],  W, Dim_Employee[Utilisation Scope] = \"Delivery\" )\n)", "0.0%",
  "Rolling 13 weeks. Trend this, not a single fortnight."),
 ("Utilisation % FYTD",
  "VAR FYc = MAX ( Dim_Date[FY] )\nVAR LastD = MAX ( Dim_Date[Date] )\n"
  "RETURN CALCULATE ( [Utilisation %], REMOVEFILTERS ( Dim_Date ), Dim_Date[FY] = FYc, Dim_Date[Date] <= LastD )",
  "0.0%", "Australian financial year to date, written without TOTALYTD's locale-sensitive year-end argument."),

 # ---- WIP hours ----
 ("WIP Hours Added",
  "CALCULATE ( [Chargeable Hours], Dim_Job[Billing Basis Final] = \"Project\" )", "#,0.0",
  "The monthly movement. Chargeable hours on date-coded project jobs. Recurring onsite contracts are billed monthly and never accrue."),
 ("WIP Hours to Date",
  "CALCULATE (\n    [Chargeable Hours],\n    Dim_Job[Billing Basis Final] = \"Project\",\n"
  "    REMOVEFILTERS ( Dim_Date ),\n    Dim_Date[Date] <= MAX ( Dim_Date[Date] )\n)", "#,0.0",
  "Cumulative chargeable hours booked to open project jobs. NOT unbilled hours until Last Fully Billed Date is populated - label it 'hours booked to date'."),
 ("Unbilled WIP Hours",
  "SUMX (\n    VALUES ( Dim_Job[Job No] ),\n"
  "    VAR LastBilled = CALCULATE ( MAX ( Dim_Job[Last Fully Billed Date] ) )\n"
  "    VAR Basis      = CALCULATE ( SELECTEDVALUE ( Dim_Job[Billing Basis Final] ) )\n"
  "    VAR UpTo       = MAX ( Dim_Date[Date] )\n"
  "    RETURN\n        IF ( Basis <> \"Project\", 0,\n"
  "            CALCULATE ( [Chargeable Hours], REMOVEFILTERS ( Dim_Date ),\n"
  "                Dim_Date[Date] <= UpTo,\n"
  "                Dim_Date[Date] > IF ( ISBLANK ( LastBilled ), DATE(1900,1,1), LastBilled ) ) )\n)",
  "#,0.0",
  "The true balance: chargeable hours since the job was last fully billed. Populate Last Fully Billed Date in map_jobs.csv from the month the job's dollar balance on the WIP schedule returns to zero."),

 # ---- data quality ----
 ("Hours With No Job Number",
  "CALCULATE ( [Total Hours], ISBLANK ( Fact_Timesheet[Job No] ) )", "#,0.0",
  "Should be zero. Timesheet lines that cannot be attributed to any job."),
 ("Unmatched Employee Hours",
  "CALCULATE ( [Total Hours], ISBLANK ( Fact_Timesheet[Employee Key] ) )", "#,0.0",
  "Should be zero. Hours whose Employee Id / Name pair is not in map_employees.csv."),
 ("Unmapped Job Hours",
  "SUMX ( FILTER ( Fact_Timesheet, NOT ISBLANK ( Fact_Timesheet[Job No] ) "
  "&& ISBLANK ( RELATED ( Dim_Job[Job No] ) ) ), Fact_Timesheet[Hours] )", "#,0.0",
  "Should be zero. A job was used on a timesheet but is missing from map_jobs.csv."),
 ("Unmapped Pay Type Hours",
  "SUMX ( FILTER ( Fact_Timesheet, ISBLANK ( RELATED ( Dim_PayType[Pay Type Key] ) ) ), Fact_Timesheet[Hours] )",
  "#,0.0", "Should be zero."),
 ("Timesheet Coverage %",
  "DIVIDE ( [Worked Hours] + [Leave Hours], [Permanent Capacity Hours] )", "0.0%",
  "Materially below 100% means missing timesheets, not low utilisation. Answer this before drawing any conclusion from Utilisation %."),
 ("Last Timesheet Date", "MAX ( Fact_Timesheet[Work Date] )", None,
  "Read this before the utilisation number. A partial month against a full-month denominator is what made August 2026 read 50.2% when July read 72.6%."),
 ("Data Age (days)",
  "VAR L = [Last Timesheet Date]\nRETURN IF ( ISBLANK ( L ), BLANK (), INT ( TODAY () - L ) )", "#,0", ""),
 ("Working Days With Data",
  "CALCULATE ( DISTINCTCOUNT ( Fact_Timesheet[Date Key] ), Dim_Date[Is Weekday] = TRUE () )", "#,0",
  "Compare against Working Days In Period. If it is lower, you are looking at a partial month."),
 ("Working Days In Period",
  "CALCULATE ( COUNTROWS ( Dim_Date ), Dim_Date[Is Weekday] = TRUE () )", "#,0", ""),
 ("Data Quality Flag",
  "VAR I =\n    IF ( [Hours With No Job Number] > 0, 1, 0 )\n"
  "  + IF ( [Unmatched Employee Hours] > 0, 1, 0 )\n"
  "  + IF ( [Unmapped Job Hours] > 0, 1, 0 )\n"
  "  + IF ( [Working Days With Data] < [Working Days In Period], 1, 0 )\n"
  "RETURN IF ( I = 0, \"OK\", I & \" issue(s) - see Data Quality page\" )", None, ""),
]

# ---------------------------------------------------------------- columns
S, I, D, DT, B = "string", "int64", "double", "dateTime", "boolean"
COLUMNS = {
 "Dim_Date": [("Date",DT),("Date Key",I),("Year",I),("Month No",I),("Month",S),("Month Sort",I),
              ("Month Year",S),("FY",I),("FY Label",S),("FY Month No",I),("FY Quarter",S),
              ("Week Ending",DT),("Day Name",S),("Day of Week No",I),("Is Weekday",B)],
 "Dim_Employee": [("Employee Key",S),("Employee Id",S),("Employee Name",S),("Person Group",S),
                  ("Employment Type",S),("First Seen",DT),("Last Seen",DT),("Rows",I),("Hours",D),
                  ("Utilisation Scope",S),("Standard Weekly Hours",D),("FTE",D),
                  ("Target Utilisation %",D),("NEEDS REVIEW",S),("Is Casual",B)],
 "Dim_Job": [("Job No",I),("Job Name",S),("Chargeable",S),("Billing Basis",S),("Job Group",S),
             ("First Booked",DT),("Last Booked",DT),("Total Hours",D),("Last Fully Billed Date",DT),
             ("Override Billing Basis",S),("Billing Basis Final",S)],
 "Dim_PayType": [("Pay Type",S),("Pay Type Trimmed",S),("Category",S),("Counts Toward Capacity",S),
                 ("Rows",I),("Hours",D),("Pay Type Key",S),("Counts Toward Capacity Flag",B)],
 "Fact_Timesheet": [("Employee Key",S),("Employee Id",S),("Employee Name",S),("Date Key",I),
                    ("Work Date",DT),("Job No",I),("Job Name Raw",S),("Pay Type",S),("Pay Type Key",S),
                    ("Chargeable Category",S),("Hours",D),("Chargeable Hours",D)],
 "Fact_Capacity": [("Employee Key",S),("Date Key",I),("Date",DT),("Capacity Hours",D)],
}
HIDDEN = {"Dim_Date":{"Date Key"}, "Dim_Employee":{"Employee Key","Rows","Hours"},
          "Dim_Job":set(), "Dim_PayType":{"Pay Type Key","Counts Toward Capacity"},
          "Fact_Timesheet":set(COLUMNS["Fact_Timesheet"][i][0] for i in range(len(COLUMNS["Fact_Timesheet"]))),
          "Fact_Capacity":set(c for c,_ in COLUMNS["Fact_Capacity"])}
SORT_BY = {"Dim_Date": {"Month":"Month Sort", "Day Name":"Day of Week No"}}

def column(tbl, name, dt):
    c = {"name": name, "dataType": dt, "sourceColumn": name,
         "summarizeBy": "sum" if dt in (D,) and "Hours" in name else "none",
         "annotations": [{"name": "SummarizationSetBy", "value": "Automatic"}]}
    if name in HIDDEN.get(tbl, set()): c["isHidden"] = True
    sb = SORT_BY.get(tbl, {}).get(name)
    if sb: c["sortByColumn"] = sb
    if tbl == "Dim_Date" and name == "Date": c["isKey"] = True
    return c

def table(name):
    t = {"name": name,
         "columns": [column(name, c, d) for c, d in COLUMNS[name]],
         "partitions": [{"name": name, "mode": "import",
                         "source": {"type": "m", "expression": M[name].strip()}}]}
    if name == "Dim_Date": t["dataCategory"] = "Time"
    if name.startswith("Fact_"): t["isHidden"] = False
    return t

measures_tbl = {
 "name": "_Measures",
 "columns": [{"name": "placeholder", "dataType": I, "sourceColumn": "placeholder", "isHidden": True,
              "summarizeBy": "none", "annotations": [{"name": "SummarizationSetBy", "value": "Automatic"}]}],
 "partitions": [{"name": "_Measures", "mode": "import",
                 "source": {"type": "m", "expression": 'let Source = #table({"placeholder"},{{1}}) in Source'}}],
 "measures": []}
for nm, expr, fmt, desc in MEASURES:
    m = {"name": nm, "expression": expr}
    if fmt: m["formatString"] = fmt
    if desc: m["description"] = desc
    measures_tbl["measures"].append(m)

def rel(ft, fc, tt, tc):
    return {"name": str(uuid.uuid5(uuid.NAMESPACE_DNS, f"{ft}.{fc}->{tt}.{tc}")),
            "fromTable": ft, "fromColumn": fc, "toTable": tt, "toColumn": tc}

RELATIONSHIPS = [
 rel("Fact_Timesheet","Date Key","Dim_Date","Date Key"),
 rel("Fact_Timesheet","Employee Key","Dim_Employee","Employee Key"),
 rel("Fact_Timesheet","Job No","Dim_Job","Job No"),
 rel("Fact_Timesheet","Pay Type Key","Dim_PayType","Pay Type Key"),
 rel("Fact_Capacity","Date Key","Dim_Date","Date Key"),
 rel("Fact_Capacity","Employee Key","Dim_Employee","Employee Key"),
]

def param(name, value, ptype="Text"):
    lit = '"%s"' % value.replace('\\', '\\\\').replace('"', '\\"') if ptype == "Text" else str(value)
    return {"name": name, "kind": "m",
            "expression": '%s meta [IsParameterQuery=true, Type="%s", IsParameterQueryRequired=true]' % (lit, ptype),
            "annotations": [{"name": "PBI_NavigationStepName", "value": "Navigation"},
                            {"name": "PBI_ResultType", "value": ptype}]}

MODEL = {
 "name": "Model",
 "compatibilityLevel": 1550,
 "model": {
   "culture": "en-AU",
   "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
   "defaultPowerBIDataSourceVersion": "powerBI_V3",
   "sourceQueryCulture": "en-AU",
   "tables": [table(t) for t in
              ["Dim_Date","Dim_Employee","Dim_Job","Dim_PayType","Fact_Timesheet","Fact_Capacity"]]
             + [measures_tbl],
   "relationships": RELATIONSHIPS,
   "expressions": [
      param("UtilWorkbookPath", r"C:\Utilisation\August_2026_Utilisation_Report.xlsx"),
      param("MappingFolderPath", r"C:\Utilisation\mapping"),
      param("FY_Start_Year", 2024, "Number"),
      param("FY_End_Year", 2027, "Number"),
   ],
   "annotations": [{"name": "__PBI_TimeIntelligenceEnabled", "value": "0"},
                   {"name": "PBIDesktopVersion", "value": "2.140.0.0"}]
 }
}

LAYOUT = {
 "id": 0, "resourcePackages": [], "config": json.dumps({"version": "5.43"}),
 "layoutOptimization": 0, "publicCustomVisuals": [],
 "sections": [{"id": 0, "name": "ReportSection", "displayName": n, "filters": "[]",
               "ordinal": i, "visualContainers": [],
               "config": "{}", "displayOption": 1, "height": 720.0, "width": 1280.0}
              for i, n in enumerate(["Summary", "By person", "Where the time went",
                                     "WIP hours", "Data quality"])]
}
for i, sec in enumerate(LAYOUT["sections"]):
    sec["name"] = "ReportSection%d" % i

CONTENT_TYPES = ('<?xml version="1.0" encoding="utf-8"?>'
 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
 '<Default Extension="json" ContentType="" />'
 '<Override PartName="/Version" ContentType="" />'
 '<Override PartName="/Settings" ContentType="" />'
 '<Override PartName="/Metadata" ContentType="" />'
 '<Override PartName="/DataModelSchema" ContentType="" />'
 '<Override PartName="/Report/Layout" ContentType="" />'
 '<Override PartName="/DiagramLayout" ContentType="" />'
 '</Types>')

def u16(o):
    return (json.dumps(o, ensure_ascii=False) if not isinstance(o, str) else o).encode("utf-16-le")

with zipfile.ZipFile(PBIT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("Version", u16("3.0"))
    z.writestr("DataModelSchema", u16(MODEL))
    z.writestr("DiagramLayout", u16({"version": 0, "diagrams": []}))
    z.writestr("Report/Layout", u16(LAYOUT))
    z.writestr("Settings", u16({"Version": 3}))
    z.writestr("Metadata", u16({"Version": 3, "AutoCreatedRelationships": [],
                                "FileDescription": "Utilisation and WIP hours model"}))
    z.writestr("SecurityBindings", b"")
    z.writestr("[Content_Types].xml", CONTENT_TYPES.encode("utf-8"))

print("built:", PBIT, os.path.getsize(PBIT), "bytes")
print("tables:", len(MODEL["model"]["tables"]),
      "| measures:", len(measures_tbl["measures"]),
      "| relationships:", len(RELATIONSHIPS),
      "| parameters:", len(MODEL["model"]["expressions"]))
