"""Replace the GL paste with a fresh Xero export.

Danica supplied Account Transactions for 1 July to 31 October 2026. It is the
same twelve columns as the P&L Analysis tab, in the same order.

Rows dated before 1 July are kept from the old paste. Those are June, which
belongs to FY2026 and is not part of this year's reporting, and the export
does not cover them.

Everything from 1 July on is replaced wholesale, so the two correcting lines
added by patch_v100 go with it. They are no longer needed.

October comes in with the export. The Summary only counts months up to the
reporting month, so October sits in its own column and stays out of the year
to date.
"""
import datetime
import sys
import openpyxl

CUT = datetime.date(2026, 7, 1)
EPOCH = datetime.date(1899, 12, 30)
COLS = 12                      # A to L, which is all the tab uses


def as_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, (int, float)):
        return EPOCH + datetime.timedelta(days=int(v))
    return None


def main(pack, export):
    src = openpyxl.load_workbook(export, data_only=True, read_only=True)
    rows = list(src['Account Transactions'].iter_rows(min_row=6, values_only=True))
    fresh = [r for r in rows if r[0] is not None]
    dates = [d for d in (as_date(r[3]) for r in fresh) if d]
    print('export: %d rows, %s to %s' % (len(fresh), min(dates), max(dates)))

    wb = openpyxl.load_workbook(pack)
    ws = wb['P&L Analysis']

    keep = []
    last = ws.max_row
    for r in range(2, last + 1):
        if ws.cell(r, 1).value is None:
            continue
        d = as_date(ws.cell(r, 4).value)
        if d is not None and d < CUT:
            keep.append([ws.cell(r, c).value for c in range(1, COLS + 1)])
    print('kept from the old paste, before 1 July: %d rows' % len(keep))

    style = [ws.cell(2, c)._style for c in range(1, COLS + 1)]

    for r in range(2, last + 1):
        for c in range(1, 14):
            ws.cell(r, c).value = None

    out = 2
    for row in keep:
        for c, v in enumerate(row, 1):
            cell = ws.cell(out, c)
            cell.value = v
            cell._style = style[c - 1]
        out += 1
    for row in fresh:
        for c in range(1, COLS + 1):
            v = row[c - 1]
            if c == 4:
                d = as_date(v)
                v = (d - EPOCH).days if d else None
            cell = ws.cell(out, c)
            cell.value = v
            cell._style = style[c - 1]
        out += 1

    print('tab now holds %d data rows' % (out - 2))
    wb.save(pack)
    print('saved')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
