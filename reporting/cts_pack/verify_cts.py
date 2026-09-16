"""Load the real Xero GL export into the pack and tie it back to the numbers
CTS already produce in PL_Analysis.xlsm and 2026-08_Graph.xlsx."""
import shutil, subprocess, sys, re
from pathlib import Path
from collections import defaultdict
from openpyxl import load_workbook

HERE = Path(__file__).parent
SRC = HERE / "CTS Financial Controller Pack.xlsx"
TEST = HERE / "_verify_cts.xlsx"
UP = Path("/root/.claude/uploads/a00ffe59-433c-507b-9006-917d8c60083a")
PL = UP / "8084a881-PL_Analysis_-_Aug_2026_-_Xero.xlsm"
RECALC = str(next(Path("/root/.claude/skills/synced").glob("*/xlsx/scripts/recalc.py")))

# ---- read their GL export -------------------------------------------------
src = load_workbook(PL, data_only=True, read_only=True)["1. Xero GL Transactions"]
rows = [r for r in src.iter_rows(min_row=2, max_col=18, values_only=True) if r[1] is not None]
print(f"GL lines from their export: {len(rows):,}")

# their own answers, straight from the same file
exp = defaultdict(float)
for r in rows:
    exp[(r[15], r[12], r[17])] += (r[13] or 0)      # Month, CostCentre, AccountType
DEPTS = ["Onsite", "Production", "Video", "Integration", "Consulting", "CTS"]

shutil.copy(SRC, TEST)
wb = load_workbook(TEST)
gl, st = wb["GL_Paste"], wb["Setup"]
st["B5"] = "Aug 2026"
for i, row in enumerate(rows):
    for j, v in enumerate(row):
        if v is not None:
            # their export carries #N/A in the Index column; never write it back as a formula
            if isinstance(v, str) and v.startswith("="):
                v = None
            gl.cell(2 + i, 1 + j, v)
wb.save(TEST)
print("pasted, recalculating (this grid is large - allow a few minutes)")
res = subprocess.run([sys.executable, RECALC, str(TEST), "1500"], capture_output=True, text=True)
print(res.stdout.strip()[:400])

vb = load_workbook(TEST, data_only=True)
fs, ct, se = vb["Financial_Summary"], vb["Controls"], vb["Setup"]
fails = 0
def chk(label, got, want, tol=0.5):
    global fails
    ok = abs((got or 0) - want) < tol
    fails += not ok
    print(f"{'OK ' if ok else 'BAD'} {label:44} theirs {want:>13,.2f}   pack {(got or 0):>13,.2f}")

print(f"\nSetup: FY {se['B7'].value}  period {se['B8'].value}  ({se['B5'].value})")
print(f"Lines pasted {se['B21'].value:,}  kept {se['B23'].value:,}  "
      f"no account {se['B26'].value}  no dept {se['B27'].value}  outside FY {se['B28'].value}")

# locate the GM-by-department block
gm0 = None
for r in range(1, fs.max_row + 1):
    if str(fs.cell(r, 1).value).strip() == "Department" and fs.cell(r, 2).value:
        gm0 = r + 1; break
print("\n--- revenue by department, AUGUST (their tab 4 / graph 'PL' month block) ---")
for i, d in enumerate(DEPTS):
    want = exp.get(("Aug", d.upper(), "Income"), 0.0)
    chk(f"{d} revenue - Aug", fs.cell(gm0 + i, 2).value, want)
print("\n--- revenue by department, YEAR TO DATE (their tab 6 / graph 'PL' YTD block) ---")
for i, d in enumerate(DEPTS):
    want = sum(exp.get((m, d.upper(), "Income"), 0.0) for m in ("Jul", "Aug"))
    chk(f"{d} revenue - YTD", fs.cell(gm0 + i, 6).value, want)
print("\n--- cost of sales by department, YEAR TO DATE ---")
for i, d in enumerate(DEPTS):
    want = sum(exp.get((m, d.upper(), "Cost of Sales"), 0.0) for m in ("Jul", "Aug"))
    chk(f"{d} COGS - YTD", fs.cell(gm0 + i, 7).value, want)

tot_inc = sum(v for (m, c, t), v in exp.items() if t == "Income")
tot_cogs = sum(v for (m, c, t), v in exp.items() if t == "Cost of Sales")
print("\n--- company totals ---")
inc_r = next(r for r in range(1, fs.max_row) if str(fs.cell(r, 1).value).strip() == "Income")
cogs_r = next(r for r in range(1, fs.max_row) if str(fs.cell(r, 1).value).strip() == "Cost of Sales")
gp_r = next(r for r in range(1, fs.max_row) if str(fs.cell(r, 1).value).strip() == "Gross Profit")
chk("company income YTD", fs.cell(inc_r, 6).value, tot_inc)
chk("company cost of sales YTD", fs.cell(cogs_r, 6).value, tot_cogs)
chk("company gross profit YTD", fs.cell(gp_r, 6).value, tot_inc + tot_cogs)
gm = fs.cell(gp_r + 1, 6).value
print(f"    gross margin YTD: {gm:.1%}" if isinstance(gm, float) else f"    gross margin: {gm}")

print("\n--- Controls ---")
for r in range(6, 25):
    if ct.cell(r, 1).value:
        s = ct.cell(r, 5).value
        fails += s == "FAIL"
        print(f"{'!! ' if s=='FAIL' else '   '}{ct.cell(r,1).value:5}{str(s):8}"
              f"{str(ct.cell(r,3).value):>14}  {str(ct.cell(r,2).value)[:52]}")
print("\nRESULT:", "TIES TO THEIR OWN NUMBERS" if fails == 0 else f"{fails} MISMATCHES")
