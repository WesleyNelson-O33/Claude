// Fact_Timesheet
// Paste into: Home > Transform data > New Source > Blank Query >
// Advanced Editor. Name the query EXACTLY Fact_Timesheet - the queries reference
// each other by name.

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
