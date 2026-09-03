// Fact_Capacity
// Paste into: Home > Transform data > New Source > Blank Query >
// Advanced Editor. Name the query EXACTLY Fact_Capacity - the queries reference
// each other by name.

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
