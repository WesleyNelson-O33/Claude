// ============================================================
// Dim_Date  -  Australian financial year calendar (1 Jul - 30 Jun)
// Mark this table as the Date table in Power BI:
//   Table tools > Mark as date table > Date column = [Date]
// ============================================================
let
    Start       = #date( FY_Start_Year, 7, 1 ),
    End         = #date( FY_End_Year + 1, 6, 30 ),
    DayCount    = Duration.Days( End - Start ) + 1,
    Dates       = List.Dates( Start, DayCount, #duration( 1, 0, 0, 0 ) ),
    T           = Table.FromList( Dates, Splitter.SplitByNothing(), { "Date" } ),
    Typed       = Table.TransformColumnTypes( T, { { "Date", type date } } ),

    A1  = Table.AddColumn( Typed, "Date Key",      each Date.Year([Date])*10000 + Date.Month([Date])*100 + Date.Day([Date]), Int64.Type ),
    A2  = Table.AddColumn( A1,    "Year",          each Date.Year([Date]), Int64.Type ),
    A3  = Table.AddColumn( A2,    "Month No",      each Date.Month([Date]), Int64.Type ),
    A4  = Table.AddColumn( A3,    "Month",         each Date.ToText([Date], [Format="MMM", Culture="en-AU"]), type text ),
    A5  = Table.AddColumn( A4,    "Month Sort",    each Date.Year([Date])*100 + Date.Month([Date]), Int64.Type ),
    A6  = Table.AddColumn( A5,    "Month Year",    each Date.ToText([Date], [Format="MMM yyyy", Culture="en-AU"]), type text ),

    // Australian FY is named by the year it ENDS: Jul 2026 - Jun 2027 = FY2027
    A7  = Table.AddColumn( A6,    "FY",            each if Date.Month([Date]) >= 7 then Date.Year([Date]) + 1 else Date.Year([Date]), Int64.Type ),
    A8  = Table.AddColumn( A7,    "FY Label",      each "FY" & Text.From( [FY] ), type text ),
    A9  = Table.AddColumn( A8,    "FY Month No",   each if Date.Month([Date]) >= 7 then Date.Month([Date]) - 6 else Date.Month([Date]) + 6, Int64.Type ),
    A10 = Table.AddColumn( A9,    "FY Quarter",    each "Q" & Text.From( Number.RoundUp( [FY Month No] / 3 ) ), type text ),

    // Payroll weeks run Monday to Sunday
    A11 = Table.AddColumn( A10,   "Week Ending",   each Date.EndOfWeek( [Date], Day.Monday ), type date ),
    A12 = Table.AddColumn( A11,   "Week Ending Key", each Date.Year([Week Ending])*10000 + Date.Month([Week Ending])*100 + Date.Day([Week Ending]), Int64.Type ),
    A13 = Table.AddColumn( A12,   "Day Name",      each Date.ToText([Date], [Format="ddd", Culture="en-AU"]), type text ),
    A14 = Table.AddColumn( A13,   "Day of Week No",each Date.DayOfWeek( [Date], Day.Monday ) + 1, Int64.Type ),
    A15 = Table.AddColumn( A14,   "Is Weekday",    each Date.DayOfWeek( [Date], Day.Monday ) < 5, type logical ),
    A16 = Table.AddColumn( A15,   "Is Past",       each [Date] <= Date.From( DateTime.LocalNow() ), type logical )
in
    A16
