"""Copies every typed (unlocked) cell from an old copy of the tracker into a new build.

Finance rows are matched by Row ID, so the new row order does not matter. Every other
sheet has the same layout, so cells are copied position for position. Formulas are never
copied - the new build's formulas stay.

Usage: python migrate_inputs.py <new_build.xlsx> <old_filled.xlsx> <out.xlsx>
"""
import sys

import openpyxl
from openpyxl.cell.cell import MergedCell

NEW, OLD, OUT = sys.argv[1:4]
new = openpyxl.load_workbook(NEW)
old = openpyxl.load_workbook(OLD)
copied = {}


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


for ws in new.worksheets:
    if ws.title not in old.sheetnames or ws.title == "Example Guide":
        continue
    src = old[ws.title]
    n = 0
    if ws.title == "Finance":
        where = {src[f"A{r}"].value: r for r in range(7, src.max_row + 1) if src[f"A{r}"].value}
        for r in range(7, ws.max_row + 1):
            sr = where.get(ws[f"A{r}"].value)
            if sr is None:
                continue
            for c in ws[r]:
                if isinstance(c, MergedCell) or c.protection.locked:
                    continue
                v = src.cell(sr, c.column).value
                if v is not None and not is_formula(v):
                    c.value = v
                    n += 1
    else:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for c in row:
                if isinstance(c, MergedCell) or c.protection.locked:
                    continue
                v = src.cell(c.row, c.column).value
                if is_formula(v):
                    continue
                if v != c.value:
                    c.value = v
                    n += 1
    copied[ws.title] = n
new.save(OUT)
print(copied)
