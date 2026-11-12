"""Give every row on a sheet a height, so the row grouping survives the save.

A row with nothing on it and no formatting is left out of the file
altogether. Excel then reads a gap in the run of outline levels and the
plus and minus buttons stop collapsing anything. The Summary sheet, which
does collapse properly, carries an explicit height on all 233 of its rows.

This walks the named sheets, gives every row up to the last used one a
height, and leaves the outline level it already has alone. It also clears
any outlinePr off the sheet so it falls back to the same default the
Summary sheet uses.

    tools/row_heights.py <workbook> <sheet name> [<sheet name> ...]
"""
import sys
import openpyxl

NORMAL = 14.5


def main():
    if len(sys.argv) < 3:
        print(__doc__.strip())
        return 1
    path, wanted = sys.argv[1], sys.argv[2:]
    wb = openpyxl.load_workbook(path)
    missing = [n for n in wanted if n not in wb.sheetnames]
    if missing:
        print('STOPPED: no such sheet: %s' % ', '.join(missing))
        return 1
    for name in wanted:
        ws = wb[name]
        ws.sheet_properties.outlinePr.summaryBelow = None
        ws.sheet_properties.outlinePr.summaryRight = None
        top = 0
        for r in range(1, (ws.max_row or 1) + 1):
            dim = ws.row_dimensions.get(r)
            lvl = (dim.outlineLevel or 0) if dim else 0
            top = max(top, lvl)
            ws.row_dimensions[r].outlineLevel = lvl
            if not (dim and dim.height):
                ws.row_dimensions[r].height = NORMAL
        ws.sheet_format.outlineLevelRow = top
        print('  %r: %d rows given a height, deepest outline level %d'
              % (name, ws.max_row or 0, top))
    wb.save(path)
    return 0


if __name__ == '__main__':
    sys.exit(main())
