// ============================================================
// Fact_Timesheet  -  from a folder of Employment Hero CSV exports
// ------------------------------------------------------------
// Grain: one row per timesheet line (employee x date x work type x
//        location x pay category).
//
// BEFORE YOU RUN THIS: open one of your actual exports, look at the
// header row, and fix the RenameMap below. Employment Hero's export
// column names differ between Payroll classic, the newer Payroll, and
// Time & Attendance exports, so there is no single correct list.
// Anything in RenameMap that does not exist in your file is skipped
// silently, so leaving extra entries in is harmless.
// ============================================================
let
    Source        = Folder.Files( TimesheetFolderPath ),

    OnlyCsv       = Table.SelectRows( Source, each Text.Lower( [Extension] ) = ".csv" ),

    Parsed        = Table.AddColumn( OnlyCsv, "Parsed", each
                        Table.PromoteHeaders(
                            Csv.Document( [Content],
                                [ Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv ] ),
                            [ PromoteAllScalars = true ] ) ),

    // (f) is the outer file row; the inner "each" binds to the CSV row,
    // so f[Name] is used explicitly to stamp the file name on every line.
    Tagged        = Table.AddColumn( Parsed, "Tagged", (f) =>
                        Table.AddColumn( f[Parsed], "Source File", each f[Name], type text ) ),

    Combined      = Table.Combine( Tagged[Tagged] ),

    // ---- rename incoming headers to model names ----
    RenameMap     = {
        { "Employee",              "Employee Name"        },
        { "Employee Name",         "Employee Name"        },
        { "Employee Id",           "Employee External ID" },
        { "Employee External Id",  "Employee External ID" },
        { "External Id",           "Employee External ID" },
        { "Date",                  "Work Date"            },
        { "Timesheet Date",        "Work Date"            },
        { "Start",                 "Start Time"           },
        { "Start Time",            "Start Time"           },
        { "End",                   "End Time"             },
        { "End Time",              "End Time"             },
        { "Break",                 "Break Minutes"        },
        { "Breaks",                "Break Minutes"        },
        { "Units",                 "Hours"                },
        { "Total Hours",           "Hours"                },
        { "Duration",              "Hours"                },
        { "Work Type",             "Work Type"            },
        { "Location",              "Location"             },
        { "Pay Category",          "Pay Category"         },
        { "Classification",        "Classification"       },
        { "Status",                "Status"               },
        { "Comments",              "Comments"             }
    },
    HeadersNow    = Table.ColumnNames( Combined ),
    Applicable    = List.Select( RenameMap, each List.Contains( HeadersNow, _{0} ) and _{0} <> _{1} ),
    // If your export happens to contain two source columns that map to the
    // same target (e.g. both "Units" and "Total Hours"), keep only the first.
    // Renaming both would throw a duplicate-column error.
    Deduped1      = List.Accumulate( Applicable, { }, ( keep, r ) =>
                        if List.Contains( List.Transform( keep, each _{1} ), r{1} )
                        then keep else keep & { r } ),
    Renamed       = Table.RenameColumns( Combined, Deduped1, MissingField.Ignore ),

    // ---- make sure every expected column exists ----
    Expected      = { "Employee Name", "Employee External ID", "Work Date", "Start Time",
                      "End Time", "Break Minutes", "Hours", "Work Type", "Location",
                      "Pay Category", "Status", "Source File" },
    Missing       = List.Difference( Expected, Table.ColumnNames( Renamed ) ),
    Padded        = List.Accumulate( Missing, Renamed,
                        (state, col) => Table.AddColumn( state, col, each null ) ),

    // ---- typing ----
    Typed         = Table.TransformColumnTypes( Padded, {
                        { "Employee Name",        type text },
                        { "Employee External ID", type text },
                        { "Work Date",            type date },
                        { "Break Minutes",        type number },
                        { "Hours",                type number },
                        { "Work Type",            type text },
                        { "Location",             type text },
                        { "Pay Category",         type text },
                        { "Status",               type text }
                    }, "en-AU" ),

    // ---- derive Hours from start/end when the export has no total ----
    // Start/End may come through as text ("9:00 AM"), as a time, or as a full
    // datetime depending on the export. "try ... otherwise" keeps a bad row
    // from failing the whole refresh - it lands as 0 hours and shows up on the
    // Data Quality page as a coverage gap rather than a broken dataset.
    WithHours     = Table.AddColumn( Typed, "Hours Final", each
                        if [Hours] <> null and [Hours] > 0 then [Hours]
                        else if [Start Time] <> null and [End Time] <> null then
                            ( try
                                let
                                    s   = Time.From( [Start Time] ),
                                    e   = Time.From( [End Time] ),
                                    raw = if e >= s then Duration.TotalHours( e - s )
                                                    else Duration.TotalHours( e - s ) + 24,
                                    brk = ( if [Break Minutes] = null then 0 else [Break Minutes] ) / 60
                                in
                                    Number.Round( raw - brk, 4 )
                              otherwise 0 )
                        else 0, type number ),

    // ---- keep only approved, paid time ----
    // Adjust this list to match the status values your export actually
    // produces. If your export has no Status column at all, this step
    // lets everything through (Status will be null).
    ApprovedOnly  = Table.SelectRows( WithHours, each
                        [Status] = null
                        or List.Contains( { "approved", "processed", "submitted and approved" },
                                          Text.Lower( Text.Trim( [Status] ) ) ) ),

    NonZero       = Table.SelectRows( ApprovedOnly, each [Hours Final] <> null and [Hours Final] > 0 ),

    // ---- keys ----
    WithKeys      = Table.AddColumn( NonZero, "Date Key", each
                        Date.Year([Work Date]) * 10000
                        + Date.Month([Work Date]) * 100
                        + Date.Day([Work Date]), Int64.Type ),

    WithEmpKey    = Table.AddColumn( WithKeys, "Employee Key", each
                        if [Employee External ID] <> null and Text.Trim([Employee External ID]) <> ""
                        then Text.Upper( Text.Trim( [Employee External ID] ) )
                        else Text.Upper( Text.Trim( [Employee Name] ) ), type text ),

    WithWtKey     = Table.AddColumn( WithEmpKey, "Work Type Key", each
                        if [Work Type] = null or Text.Trim([Work Type]) = ""
                        then "UNSPECIFIED"
                        else Text.Upper( Text.Trim( [Work Type] ) ), type text ),

    // ---- de-duplicate re-exported periods ----
    // If you re-export the same fortnight, the same line appears twice.
    // This drops exact duplicates while keeping genuinely repeated
    // lines apart by their source row content.
    Deduped       = Table.Distinct( WithWtKey, { "Employee Key", "Work Date", "Work Type Key",
                                                 "Location", "Start Time", "End Time", "Hours Final" } ),

    Final         = Table.SelectColumns( Deduped, {
                        "Employee Key", "Employee Name", "Date Key", "Work Date",
                        "Work Type Key", "Work Type", "Location", "Pay Category",
                        "Hours Final", "Status", "Source File" } ),

    Renamed2      = Table.RenameColumns( Final, { { "Hours Final", "Hours" } } )
in
    Renamed2
