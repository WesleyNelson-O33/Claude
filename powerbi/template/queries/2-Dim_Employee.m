// Dim_Employee
// Paste into: Home > Transform data > New Source > Blank Query >
// Advanced Editor. Name the query EXACTLY Dim_Employee - the queries reference
// each other by name.

let
    Src      = Csv.Document(File.Contents(Text.TrimEnd(MappingFolderPath,"\") & "\map_employees.csv"),
                            [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Head     = Table.PromoteHeaders(Src, [PromoteAllScalars=true]),
    Typed    = Table.TransformColumnTypes(Head, {
                  {"Employee Key", type text}, {"Employee Id", type text}, {"Employee Name", type text},
                  {"Person Group", type text}, {"Employment Type", type text},
                  {"First Seen", type date}, {"Last Seen", type date},
                  {"Utilisation Scope", type text}, {"Standard Weekly Hours", type number},
                  {"FTE", type number}, {"Target Utilisation %", type number}}, "en-AU"),
    // sensible defaults so the model works before every cell is filled in
    D1 = Table.ReplaceValue(Typed, null, "Delivery", Replacer.ReplaceValue, {"Utilisation Scope"}),
    D2 = Table.ReplaceValue(D1,    null, 38,         Replacer.ReplaceValue, {"Standard Weekly Hours"}),
    D3 = Table.ReplaceValue(D2,    null, 1,          Replacer.ReplaceValue, {"FTE"}),
    D4 = Table.ReplaceValue(D3,    null, 0,          Replacer.ReplaceValue, {"Target Utilisation %"}),
    IsCasual = Table.AddColumn(D4, "Is Casual", each Text.Lower([Employment Type]) = "casual", type logical),
    Keep     = Table.SelectRows(IsCasual, each [Employee Key] <> null and [Employee Key] <> ""),
    Unique   = Table.Distinct(Keep, {"Employee Key"})
in
    Unique
