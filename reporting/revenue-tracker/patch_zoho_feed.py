"""Adds the Zoho dump feed to an existing filled tracker (keeps every typed cell, comment and extra sheet),
optionally pastes a Zoho export into Zoho Paste exactly as a user would, and refreshes the Read Me.

Usage: python patch_zoho_feed.py <tracker.xlsx> <out.xlsx> <fresh_blank_build.xlsx> [zoho_export.xlsx]
"""
import sys
from copy import copy

import openpyxl

import zoho_feed

TR, OUT, BLANK = sys.argv[1:4]
ZOHO = sys.argv[4] if len(sys.argv) > 4 else None
wb = openpyxl.load_workbook(TR)
zoho_feed.apply(wb)

if ZOHO:        # the monthly dump: copy the export's cells into A1 onwards, as Ctrl+C / Ctrl+V would
    src = openpyxl.load_workbook(ZOHO).active
    dst = wb["Zoho Paste"]
    for row in src.iter_rows(min_row=1, max_row=src.max_row, max_col=8):
        for c in row:
            dst.cell(c.row, c.column).value = c.value

# Read Me from the fresh build
src = openpyxl.load_workbook(BLANK)["Read Me"]
dst = wb["Read Me"]
for row in dst.iter_rows():
    for c in row:
        c.value = None
for row in src.iter_rows():
    for c in row:
        d = dst.cell(c.row, c.column)
        d.value = c.value
        if c.has_style:
            d.font, d.fill, d.border = copy(c.font), copy(c.fill), copy(c.border)
            d.alignment, d.protection = copy(c.alignment), copy(c.protection)
for k, v in src.row_dimensions.items():
    dst.row_dimensions[k].height = v.height
wb.save(OUT)
print("saved", OUT)
