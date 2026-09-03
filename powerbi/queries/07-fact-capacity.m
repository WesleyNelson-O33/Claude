// ============================================================
// Fact_Capacity  -  one row per employee per working day
// ------------------------------------------------------------
// This is the DENOMINATOR of the utilisation ratio, and it does not
// exist anywhere in Employment Hero - you have to build it.
//
// ASSUMPTION BAKED IN HERE: contracted weekly hours are spread evenly
// across Monday to Friday, and a public holiday in the employee's state
// removes one standard day of capacity. If you run a rostered
// operation, replace this query with the roster export instead - see
// README section 6.
// ============================================================
let
    Emp        = Table.SelectColumns( Dim_Employee,
                    { "Employee Key", "Standard Weekly Hours", "Start Date", "End Date", "State" } ),

    Weekdays   = Table.SelectColumns(
                    Table.SelectRows( Dim_Date, each [Is Weekday] = true ),
                    { "Date", "Date Key" } ),

    Cross      = Table.AddColumn( Emp, "Cal", each Weekdays ),
    Expanded   = Table.ExpandTableColumn( Cross, "Cal", { "Date", "Date Key" } ),

    // only days the person was actually employed
    InService  = Table.SelectRows( Expanded, each
                    [Date] >= [Start Date]
                    and ( [End Date] = null or [Date] <= [End Date] ) ),

    BaseCap    = Table.AddColumn( InService, "Base Capacity Hours",
                    each ( if [Standard Weekly Hours] = null then 0 else [Standard Weekly Hours] ) / 5, type number ),

    // knock out public holidays in that employee's state
    WithPHKey  = Table.AddColumn( BaseCap, "PH Key",
                    each Text.Upper( Text.Trim( if [State] = null then "" else [State] ) ) & "|" &
                         Date.ToText( [Date], [Format="yyyy-MM-dd"] ), type text ),

    Joined     = Table.NestedJoin( WithPHKey, { "PH Key" },
                                   Dim_PublicHoliday, { "PH Key" },
                                   "PH", JoinKind.LeftOuter ),
    PHCount    = Table.AddColumn( Joined, "Is Public Holiday",
                    each Table.RowCount( [PH] ) > 0, type logical ),

    NetCap     = Table.AddColumn( PHCount, "Capacity Hours",
                    each if [Is Public Holiday] then 0 else [Base Capacity Hours], type number ),

    Final      = Table.SelectColumns( NetCap, {
                    "Employee Key", "Date Key", "Date",
                    "Capacity Hours", "Base Capacity Hours", "Is Public Holiday" } )
in
    Final
