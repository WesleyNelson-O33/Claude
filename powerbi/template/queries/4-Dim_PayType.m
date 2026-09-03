// Dim_PayType
// Paste into: Home > Transform data > New Source > Blank Query >
// Advanced Editor. Name the query EXACTLY Dim_PayType - the queries reference
// each other by name.

let
    Src   = Csv.Document(File.Contents(Text.TrimEnd(MappingFolderPath,"\") & "\map_paytypes.csv"),
                         [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Head  = Table.PromoteHeaders(Src, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Head, {
               {"Pay Type", type text}, {"Pay Type Trimmed", type text},
               {"Category", type text}, {"Counts Toward Capacity", type text}}, "en-AU"),
    Key   = Table.AddColumn(Typed, "Pay Type Key",
               each Text.Upper(Text.Combine(List.Select(Text.Split(Text.Trim([Pay Type])," "), each _ <> ""), " ")), type text),
    Flag  = Table.AddColumn(Key, "Counts Toward Capacity Flag",
               each Text.Lower(Text.Trim([Counts Toward Capacity])) = "yes", type logical),
    Uniq  = Table.Distinct(Flag, {"Pay Type Key"})
in
    Uniq
