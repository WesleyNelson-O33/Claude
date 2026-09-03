// Dim_Job
// Paste into: Home > Transform data > New Source > Blank Query >
// Advanced Editor. Name the query EXACTLY Dim_Job - the queries reference
// each other by name.

let
    Src   = Csv.Document(File.Contents(Text.TrimEnd(MappingFolderPath,"\") & "\map_jobs.csv"),
                         [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Head  = Table.PromoteHeaders(Src, [PromoteAllScalars=true]),
    Typed = Table.TransformColumnTypes(Head, {
               {"Job No", Int64.Type}, {"Job Name", type text}, {"Chargeable", type text},
               {"Billing Basis", type text}, {"Job Group", type text},
               {"First Booked", type date}, {"Last Booked", type date},
               {"Last Fully Billed Date", type date}, {"Override Billing Basis", type text}}, "en-AU"),
    // a manual override always wins over the derived basis
    Final = Table.AddColumn(Typed, "Billing Basis Final",
               each if [Override Billing Basis] <> null and Text.Trim([Override Billing Basis]) <> ""
                    then Text.Trim([Override Billing Basis]) else [Billing Basis], type text),
    Uniq  = Table.Distinct(Final, {"Job No"})
in
    Uniq
