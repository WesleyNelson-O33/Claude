// ============================================================
// Dim_WorkType  -  from mapping/worktype-utilisation-map.csv
// ------------------------------------------------------------
// THIS TABLE IS THE WHOLE REPORT. Employment Hero has no native
// "billable" flag - the utilisation number is entirely a product of how
// you classify each Work Type here. Get this wrong and every number
// downstream is wrong.
//
// The query below also appends any Work Type that appears in the
// timesheets but is MISSING from your mapping file, tagged as
// "UNMAPPED". Put a card on the report showing UNMAPPED hours so you
// notice the day someone adds a new work type in Employment Hero.
// ============================================================
let
    Source     = Csv.Document(
                    File.Contents( Text.TrimEnd( MappingFolderPath, "\" ) & "\worktype-utilisation-map.csv" ),
                    [ Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv ] ),
    Promoted   = Table.PromoteHeaders( Source, [PromoteAllScalars = true] ),

    Typed      = Table.TransformColumnTypes( Promoted, {
                    { "Work Type",             type text },
                    { "Utilisation Category",  type text },
                    { "Counts Toward Capacity",type text },
                    { "Notes",                 type text }
                 } ),

    WithKey    = Table.AddColumn( Typed, "Work Type Key",
                    each Text.Upper( Text.Trim( [Work Type] ) ), type text ),

    WithFlag   = Table.AddColumn( WithKey, "Counts Toward Capacity Flag",
                    each Text.Lower( Text.Trim( [Counts Toward Capacity] ) ) = "yes", type logical ),

    Mapped     = Table.Distinct( WithFlag, { "Work Type Key" } ),

    // ---- catch work types that exist in the data but not the map ----
    InData     = Table.Distinct(
                    Table.SelectColumns( Fact_Timesheet, { "Work Type Key", "Work Type" } ),
                    { "Work Type Key" } ),
    Unmapped   = Table.SelectRows( InData,
                    each not List.Contains( Mapped[Work Type Key], [Work Type Key] ) ),
    UnmappedT  = Table.AddColumn(
                    Table.AddColumn(
                        Table.AddColumn(
                            Table.AddColumn( Unmapped, "Utilisation Category", each "UNMAPPED", type text ),
                            "Counts Toward Capacity", each "Yes", type text ),
                        "Notes", each "Not present in worktype-utilisation-map.csv - classify it", type text ),
                    "Counts Toward Capacity Flag", each true, type logical ),

    Combined   = Table.Combine( { Mapped, UnmappedT } ),

    Ordered    = Table.SelectColumns( Combined, {
                    "Work Type Key", "Work Type", "Utilisation Category",
                    "Counts Toward Capacity", "Counts Toward Capacity Flag", "Notes" } )
in
    Ordered
