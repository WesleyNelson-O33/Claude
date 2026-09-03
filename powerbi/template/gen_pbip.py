#!/usr/bin/env python3
"""Generates the PBIP project from the same model definition as everything else."""
import importlib.util, os, json, io, re, sys, uuid, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("b", os.path.join(HERE, "build_pbit.py"))
b = importlib.util.module_from_spec(spec)
_argv = sys.argv; sys.argv = ["build_pbit.py", "/tmp"]; spec.loader.exec_module(b); sys.argv = _argv

NAME = "Utilisation"
ROOT = os.path.join(HERE, "pbip")
SM   = os.path.join(ROOT, f"{NAME}.SemanticModel")
RP   = os.path.join(ROOT, f"{NAME}.Report")
if os.path.isdir(ROOT): shutil.rmtree(ROOT)
for d in [SM, os.path.join(SM, "definition", "tables"),
          RP, os.path.join(RP, "definition", "pages")]:
    os.makedirs(d, exist_ok=True)

def w(path, text, nl="\n"):
    with io.open(path, "w", encoding="utf-8", newline=nl) as f:
        f.write(text)

def wj(path, obj):
    w(path, json.dumps(obj, indent=2, ensure_ascii=False) + "\n")

def q(n):
    """TMDL quotes any name that is not a bare identifier."""
    return n if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", n) else "'%s'" % n.replace("'", "''")

def desc(text, indent):
    if not text: return []
    return ["%s/// %s" % (indent, l) for l in text.split("\n")]

# ---------------------------------------------------------------- pointers
wj(os.path.join(ROOT, f"{NAME}.pbip"), {
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
  "version": "1.0",
  "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
  "settings": {"enableAutoRecovery": True}})

wj(os.path.join(SM, "definition.pbism"), {
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
  "version": "4.2",
  "settings": {"qnaEnabled": True}})

wj(os.path.join(RP, "definition.pbir"), {
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
  "version": "4.0",
  "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}})

for folder, kind, disp in [(SM, "SemanticModel", NAME), (RP, "Report", NAME)]:
    wj(os.path.join(folder, ".platform"), {
      "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
      "metadata": {"type": kind, "displayName": disp},
      "config": {"version": "2.0", "logicalId": str(uuid.uuid5(uuid.NAMESPACE_DNS, kind + disp))}})

# ---------------------------------------------------------------- TMDL core
DEF = os.path.join(SM, "definition")
w(os.path.join(DEF, "database.tmdl"),
  "database %s\n\tcompatibilityLevel: 1567\n" % q(NAME))

TABLES = ["Dim_Date","Dim_Employee","Dim_Job","Dim_PayType","Fact_Timesheet","Fact_Capacity","_Measures"]
model = ["model Model",
         "\tculture: en-AU",
         "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
         "\tdiscourageImplicitMeasures: true",
         "\tsourceQueryCulture: en-AU",
         ""]
for t in TABLES: model.append("ref table %s" % q(t))
model += ["", "ref relationships"] if False else [""]
w(os.path.join(DEF, "model.tmdl"), "\n".join(model) + "\n")

# ---------------------------------------------------------------- parameters
ex = []
for name, val, ptype in [("UtilWorkbookPath", r"C:\Utilisation\August 2026 Utilisation Report.xlsx", "Text"),
                         ("MappingFolderPath", r"C:\Utilisation", "Text"),
                         ("FY_Start_Year", 2024, "Number"),
                         ("FY_End_Year", 2027, "Number")]:
    lit = '"%s"' % str(val).replace("\\", "\\\\") if ptype == "Text" else str(val)
    ex.append('expression %s = %s meta [IsParameterQuery=true, Type="%s", IsParameterQueryRequired=true]'
              % (q(name), lit, ptype))
    ex.append("")
w(os.path.join(DEF, "expressions.tmdl"), "\n".join(ex))

# ---------------------------------------------------------------- relationships
rel = []
for r in b.RELATIONSHIPS:
    rname = "%s to %s" % (r["fromTable"], r["toTable"])
    rel += ["relationship %s" % q(rname),
            "\tfromColumn: %s.%s" % (q(r["fromTable"]), q(r["fromColumn"])),
            "\ttoColumn: %s.%s"   % (q(r["toTable"]),   q(r["toColumn"])),
            "\tcrossFilteringBehavior: oneDirection",
            ""]
w(os.path.join(DEF, "relationships.tmdl"), "\n".join(rel))

# ---------------------------------------------------------------- tables
def render_table(name):
    L = ["table %s" % q(name), ""]
    if name == "Dim_Date":
        L += ["\tdataCategory: Time", ""]
    # measures first
    if name == "_Measures":
        for nm, expr, fmt, dsc in b.MEASURES:
            L += desc(dsc, "\t")
            if "\n" in expr:
                # the documented form puts the block body and closing fence at
                # two tabs, level with formatString
                L.append("\tmeasure %s = ```" % q(nm))
                for line in expr.split("\n"):
                    L.append("\t\t%s" % line)
                L.append("\t\t```")
            else:
                L.append("\tmeasure %s = %s" % (q(nm), expr))
            L.append("\t\tformatString: %s" % (fmt if fmt else "0"))
            L.append("")
    # columns
    cols = b.COLUMNS[name] if name in b.COLUMNS else [("placeholder", "int64")]
    for cname, dt in cols:
        L.append("\tcolumn %s" % q(cname))
        L.append("\t\tdataType: %s" % dt)
        if name == "Dim_Date" and cname == "Date":
            L.append("\t\tisKey")
        if cname in b.HIDDEN.get(name, set()) or name == "_Measures":
            L.append("\t\tisHidden")
        L.append("\t\tsummarizeBy: %s" % ("sum" if dt == "double" and "Hours" in cname else "none"))
        L.append("\t\tsourceColumn: %s" % cname)
        sb = b.SORT_BY.get(name, {}).get(cname)
        if sb: L.append("\t\tsortByColumn: %s" % q(sb))
        L.append("")
    # partition
    L.append("\tpartition %s = m" % q(name))
    L.append("\t\tmode: import")
    L.append("\t\tsource =")
    src = b.M[name].strip() if name in b.M else 'let Source = #table({"placeholder"},{{1}}) in Source'
    for line in src.split("\n"):
        L.append("\t\t\t%s" % line if line.strip() else "")
    L.append("")
    return "\n".join(L)

for t in TABLES:
    w(os.path.join(DEF, "tables", "%s.tmdl" % t), render_table(t))

# ---------------------------------------------------------------- report
PAGES = ["Summary", "By person", "Where the time went", "WIP hours", "Data quality"]
ids = []
# PBIR requires a version.json in the definition folder. Desktop reported
# "Cannot find file 'version.json'" without it. Exact contents unverified -
# if the project still will not open, use the folder-swap route in README-FIRST.
wj(os.path.join(RP, "definition", "version.json"), {
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/versionMetadata/1.0.0/schema.json",
  "version": "1.0"})
wj(os.path.join(RP, "definition", "report.json"), {
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/1.0.0/schema.json",
  "themeCollection": {"baseTheme": {"name": "CY24SU06", "reportVersionAtImport": "5.55",
                                    "type": "SharedResources"}},
  "config": {"version": 5, "defaultDrillFilterOtherVisuals": True},
  "objects": {}})
for i, p in enumerate(PAGES):
    pid = "Page%02d" % (i + 1)
    ids.append(pid)
    pdir = os.path.join(RP, "definition", "pages", pid)
    os.makedirs(pdir, exist_ok=True)
    wj(os.path.join(pdir, "page.json"), {
      "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/1.0.0/schema.json",
      "name": pid, "displayName": p, "displayOption": "FitToPage",
      "height": 720, "width": 1280})
wj(os.path.join(RP, "definition", "pages", "pages.json"), {
  "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
  "pageOrder": ids, "activePageName": ids[0]})

w(os.path.join(ROOT, ".gitignore"), ".pbi/\n**/localSettings.json\n*.abf\n**/cache.abf\n")

n = sum(len(f) for _, _, f in os.walk(ROOT))
print("PBIP written to", ROOT, "-", n, "files")
