// ============================================================
// Fact_Timesheet  -  from the Employment Hero Payroll API
// ------------------------------------------------------------
// USE THIS ONLY IF THE CSV PATH IS NOT ENOUGH. Read the honest
// assessment in README section 3 first - for most businesses the CSV
// folder is faster to build and cheaper to run.
//
// VERIFY BEFORE USE:
//   * API access requires a Platinum plan or above.
//   * Confirm the exact endpoint path and query parameter names in the
//     current docs at https://api.keypay.com.au/australia/reference
//     (Employment Hero Payroll is the former KeyPay platform). The path
//     below is the shape the API uses, but Employment Hero has changed
//     endpoint detail before - do not assume it is right.
//   * Auth below is HTTP Basic with your API key as the username and an
//     empty password. Confirm this against your tenant.
//
// CRITICAL POWER BI GOTCHA:
//   Web.Contents must be called with a STATIC base URL and the variable
//   parts passed via RelativePath and Query. If you concatenate the URL
//   into one string, the query works in Power BI Desktop and then fails
//   on the Service with "the query references dynamic data sources".
//   That is why PayrollApiBase is a bare host and nothing else.
// ============================================================
let
    // ---- credentials -------------------------------------------------
    // Store the key in Power BI as a Basic credential (username = key,
    // password = blank) rather than hard-coding it. If you must inline it
    // for a first test, replace the line below and DELETE IT AFTERWARDS -
    // never commit a key to this repo.
    ApiKey        = "" ,   // leave blank when using the credential store
    AuthHeader    = if ApiKey = "" then []
                    else [ Authorization = "Basic " &
                           Binary.ToText( Text.ToBinary( ApiKey & ":" ), BinaryEncoding.Base64 ) ],

    FromDate      = Date.ToText( #date( FY_Start_Year, 7, 1 ), [Format="yyyy-MM-dd"] ),
    ToDate        = Date.ToText( Date.From( DateTime.LocalNow() ), [Format="yyyy-MM-dd"] ),

    // ---- paged fetch -------------------------------------------------
    PageSize      = 300,

    GetPage       = ( skip as number ) as list =>
        let
            Response = Web.Contents(
                            PayrollApiBase,
                            [
                                RelativePath = "api/v2/business/" & BusinessId & "/timesheet",
                                Query = [
                                    fromDate = FromDate,
                                    toDate   = ToDate,
                                    #"$top"  = Text.From( PageSize ),
                                    #"$skip" = Text.From( skip )
                                ],
                                Headers = AuthHeader
                            ] ),
            Json     = Json.Document( Response ),
            AsList   = if Json is list then Json
                       else if Json is record and Record.HasFields( Json, "items" ) then Json[items]
                       else { Json }
        in
            AsList,

    AllPages      = List.Generate(
                        () => [ Skip = 0, Rows = GetPage(0) ],
                        each List.Count( [Rows] ) > 0,
                        each [ Skip = [Skip] + PageSize, Rows = GetPage( [Skip] + PageSize ) ],
                        each [Rows] ),

    Flat          = List.Combine( AllPages ),

    ToTable       = Table.FromList( Flat, Splitter.SplitByNothing(), {"r"} ),

    // Expand whatever the API actually returns, then map the fields you
    // need. Run the query once, look at the expanded column list, and
    // trim this to the fields that exist.
    // Field names are gathered from the first 200 records, not just the first,
    // because optional fields are omitted from records that do not use them.
    FieldNames    = List.Distinct(
                        List.Combine(
                            List.Transform( List.FirstN( Flat, 200 ), Record.FieldNames ) ) ),
    Expanded      = if List.IsEmpty( Flat )
                    then #table( { "Employee External ID", "Employee Name", "Start Time",
                                   "End Time", "Hours", "Work Type", "Location", "Status" }, { } )
                    else Table.ExpandRecordColumn( ToTable, "r", FieldNames ),

    // ---- normalise to the same shape as the CSV query ----------------
    // Rename here so Fact_Timesheet has identical columns either way.
    Normalised    = Table.RenameColumns( Expanded, {
                        { "employeeId",   "Employee External ID" },
                        { "employeeName", "Employee Name"        },
                        { "startTime",    "Start Time"           },
                        { "endTime",      "End Time"             },
                        { "units",        "Hours"                },
                        { "workTypeName", "Work Type"            },
                        { "locationName", "Location"             },
                        { "status",       "Status"               }
                    }, MissingField.Ignore ),

    WithDate      = Table.AddColumn( Normalised, "Work Date",
                        each Date.From( [Start Time] ), type date ),

    WithKeys      = Table.AddColumn( WithDate, "Date Key", each
                        Date.Year([Work Date]) * 10000
                        + Date.Month([Work Date]) * 100
                        + Date.Day([Work Date]), Int64.Type ),

    WithEmpKey    = Table.AddColumn( WithKeys, "Employee Key",
                        each Text.Upper( Text.Trim( Text.From( [Employee External ID] ) ) ), type text ),

    WithWtKey     = Table.AddColumn( WithEmpKey, "Work Type Key", each
                        if [Work Type] = null or Text.Trim([Work Type]) = ""
                        then "UNSPECIFIED"
                        else Text.Upper( Text.Trim( [Work Type] ) ), type text ),

    AddSource     = Table.AddColumn( WithWtKey, "Source File", each "API", type text ),
    AddPayCat     = if List.Contains( Table.ColumnNames( AddSource ), "Pay Category" )
                    then AddSource
                    else Table.AddColumn( AddSource, "Pay Category", each null, type text ),

    Final         = Table.SelectColumns( AddPayCat, {
                        "Employee Key", "Employee Name", "Date Key", "Work Date",
                        "Work Type Key", "Work Type", "Location", "Pay Category",
                        "Hours", "Status", "Source File" } )
in
    Final
