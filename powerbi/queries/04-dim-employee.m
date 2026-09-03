// ============================================================
// Dim_Employee  -  from mapping/employee-targets.csv
// ------------------------------------------------------------
// This is a MANUALLY MAINTAINED file. Employment Hero will not give you
// target utilisation or a team structure, so you own that here. Export
// the employee list from Employment Hero, paste in the IDs and names,
// then fill in Team / Role / Target.
// ============================================================
let
    Source     = Csv.Document(
                    File.Contents( Text.TrimEnd( MappingFolderPath, "\" ) & "\employee-targets.csv" ),
                    [ Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv ] ),
    Promoted   = Table.PromoteHeaders( Source, [PromoteAllScalars = true] ),

    Typed      = Table.TransformColumnTypes( Promoted, {
                    { "Employee External ID",  type text   },
                    { "Employee Name",         type text   },
                    { "Team",                  type text   },
                    { "Role",                  type text   },
                    { "Standard Weekly Hours", type number },
                    { "FTE",                   type number },
                    { "Target Utilisation %",  type number },
                    { "Employment Type",       type text   },
                    { "Start Date",            type date   },
                    { "End Date",              type date   },
                    { "State",                 type text   }
                 }, "en-AU" ),

    WithKey    = Table.AddColumn( Typed, "Employee Key",
                    each Text.Upper( Text.Trim( [Employee External ID] ) ), type text ),

    // Guard: a blank or duplicated key silently breaks the relationship
    NonBlank   = Table.SelectRows( WithKey, each [Employee Key] <> null and [Employee Key] <> "" ),
    Unique     = Table.Distinct( NonBlank, { "Employee Key" } )
in
    Unique
