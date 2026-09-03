#!/usr/bin/env python3
"""Emits the guaranteed-to-work manual build: M queries + a Tabular Editor script."""
import importlib.util, os, json, re, io
spec = importlib.util.spec_from_file_location("b", os.path.join(os.path.dirname(os.path.abspath(__file__)), "build_pbit.py"))
b = importlib.util.module_from_spec(spec)
import sys; _argv = sys.argv; sys.argv = ["build_pbit.py", "/tmp"]
spec.loader.exec_module(b); sys.argv = _argv

OUT = os.path.dirname(os.path.abspath(__file__))
qd = os.path.join(OUT, "queries"); os.makedirs(qd, exist_ok=True)

ORDER = ["Dim_Date","Dim_Employee","Dim_Job","Dim_PayType","Fact_Timesheet","Fact_Capacity"]
for i, name in enumerate(ORDER, 1):
    with io.open(os.path.join(qd, f"{i}-{name}.m"), "w", encoding="utf-8") as f:
        f.write(f"// {name}\n// Paste into: Home > Transform data > New Source > Blank Query >\n"
                f"// Advanced Editor. Name the query EXACTLY {name} - the queries reference\n"
                f"// each other by name.\n\n")
        f.write(b.M[name].strip() + "\n")
print(f"wrote {len(ORDER)} M files to queries/")

# ---- Tabular Editor C# script: measures, relationships, formatting, hiding ----
def cs(s): return s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n")

lines = ['// Utilisation model - measures, relationships and formatting.',
         '// Run in Tabular Editor 2 (free): Tools > Manage BPA... no - use the C# Script tab,',
         '// paste this, press F5, then Ctrl+S to write it back to Power BI Desktop.',
         '// Requires the 6 queries to be loaded first, and Desktop left open.',
         '',
         'Table m;',
         'if (Model.Tables.Contains("_Measures")) m = Model.Tables["_Measures"];',
         'else { m = Model.AddCalculatedTable("_Measures", "{1}"); m.Columns[0].IsHidden = true; }',
         '']
for nm, expr, fmt, desc in b.MEASURES:
    v = re.sub(r'\W', '', nm)
    lines.append(f'var {v} = m.AddMeasure("{cs(nm)}", "{cs(expr)}");')
    if fmt:  lines.append(f'{v}.FormatString = "{cs(fmt)}";')
    if desc: lines.append(f'{v}.Description = "{cs(desc)}";')
lines.append('')
lines.append('// relationships')
for r in b.RELATIONSHIPS:
    lines.append(f'if (Model.Tables.Contains("{r["fromTable"]}") && Model.Tables.Contains("{r["toTable"]}"))')
    lines.append(f'    Model.AddRelationship().FromColumn = Model.Tables["{r["fromTable"]}"].Columns["{r["fromColumn"]}"];')
lines.append('')
lines.append('// NOTE: Tabular Editor cannot set both ends of a relationship in one statement.')
lines.append('// The block below does it properly.')
lines.append('var rels = new [] {')
for r in b.RELATIONSHIPS:
    lines.append(f'    new [] {{"{r["fromTable"]}","{r["fromColumn"]}","{r["toTable"]}","{r["toColumn"]}"}},')
lines.append('};')
lines.append('''foreach (var r in rels) {
    if (!Model.Tables.Contains(r[0]) || !Model.Tables.Contains(r[2])) continue;
    var ft = Model.Tables[r[0]]; var tt = Model.Tables[r[2]];
    if (!ft.Columns.Contains(r[1]) || !tt.Columns.Contains(r[3])) continue;
    var exists = Model.Relationships.Any(x => x.FromTable == ft && x.ToTable == tt
                 && x.FromColumn.Name == r[1] && x.ToColumn.Name == r[3]);
    if (exists) continue;
    var rel = Model.AddRelationship();
    rel.FromColumn = ft.Columns[r[1]];
    rel.ToColumn   = tt.Columns[r[3]];
    rel.FromCardinality = RelationshipEndCardinality.Many;
    rel.ToCardinality   = RelationshipEndCardinality.One;
    rel.CrossFilteringBehavior = CrossFilteringBehavior.OneDirection;
    rel.IsActive = true;
}''')
lines.append('')
lines.append('// sort-by columns and the date table')
lines.append('''if (Model.Tables.Contains("Dim_Date")) {
    var d = Model.Tables["Dim_Date"];
    d.DataCategory = "Time";
    if (d.Columns.Contains("Date")) ((Column)d.Columns["Date"]).IsKey = true;
    if (d.Columns.Contains("Month") && d.Columns.Contains("Month Sort"))
        d.Columns["Month"].SortByColumn = d.Columns["Month Sort"];
    if (d.Columns.Contains("Day Name") && d.Columns.Contains("Day of Week No"))
        d.Columns["Day Name"].SortByColumn = d.Columns["Day of Week No"];
}''')
lines.append('')
lines.append('// hide the plumbing so nobody drags a raw column onto a visual')
hide = []
for t, cols in b.HIDDEN.items():
    for c in sorted(cols):
        hide.append(f'    new [] {{"{t}","{c}"}},')
lines.append('var hide = new [] {')
lines.extend(hide)
lines.append('};')
lines.append('''foreach (var h in hide2()) {}
foreach (var h in hide) {
    if (Model.Tables.Contains(h[0]) && Model.Tables[h[0]].Columns.Contains(h[1]))
        Model.Tables[h[0]].Columns[h[1]].IsHidden = true;
}''')
script = "\n".join(lines)
# drop the two scaffolding lines that were only there to explain the pattern
script = script.replace('\n'.join([
 'foreach (var h in hide2()) {}',]) + '\n', '')
script = re.sub(r'// relationships\n(if \(Model\.Tables\.Contains.*\n    Model\.AddRelationship\(\).*\n)+\n// NOTE: Tabular Editor cannot set both ends of a relationship in one statement\.\n// The block below does it properly\.\n',
                '// relationships\n', script)
with io.open(os.path.join(OUT, "measures.csx"), "w", encoding="utf-8") as f:
    f.write(script + "\n")
print(f"wrote measures.csx  ({len(b.MEASURES)} measures, {len(b.RELATIONSHIPS)} relationships)")
