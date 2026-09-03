// Dim_Date
// Paste into: Home > Transform data > New Source > Blank Query >
// Advanced Editor. Name the query EXACTLY Dim_Date - the queries reference
// each other by name.

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
