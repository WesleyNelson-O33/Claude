// ============================================================
// Dim_PublicHoliday  -  from mapping/public-holidays.csv
// ------------------------------------------------------------
// Australian public holidays are state-based and change every year, and
// there is no free, stable, authoritative feed that is safe to bind a
// scheduled refresh to. Maintain this file by hand each year from
// fairwork.gov.au / your state government's list. It is 10 minutes a
// year and it is auditable, which matters if utilisation feeds a bonus.
// ============================================================
let
    Source   = Csv.Document(
                  File.Contents( Text.TrimEnd( MappingFolderPath, "\" ) & "\public-holidays.csv" ),
                  [ Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv ] ),
    Promoted = Table.PromoteHeaders( Source, [PromoteAllScalars = true] ),
    Typed    = Table.TransformColumnTypes( Promoted, {
                  { "Date",         type date   },
                  { "Holiday Name", type text   },
                  { "State",        type text   },
                  { "Hours",        type number }
               }, "en-AU" ),
    WithKey  = Table.AddColumn( Typed, "PH Key",
                  each Text.Upper( Text.Trim( [State] ) ) & "|" &
                       Date.ToText( [Date], [Format="yyyy-MM-dd"] ), type text )
in
    WithKey
