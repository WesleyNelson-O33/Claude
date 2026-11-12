"""Make the P&L FY27 snip for the deck, columns A to O, collapsed.

The sheet is 233 rows deep. Dropped on a slide whole, every row would be a
sixtieth of an inch tall and nobody could read a figure on it. So this takes
the sheet the way it looks with the groups collapsed, which is how anyone
would present it: the lines down to net profit per cent, with the account
detail folded away.

Columns A to O, as asked: the twelve months, the FY total and the YTD.

    deck/edits/snip_plfy27.py <pack.xlsx> <out.png> [max outline level]
"""
import copy
import os
import subprocess
import sys

import openpyxl
from PIL import Image, ImageChops
from openpyxl.utils import get_column_letter as L

SHEET = 'P&L FY27'
LASTCOL = 15                      # column O
LASTROW = 182                     # down to Net profit %, before the dept blocks


def main():
    if len(sys.argv) < 3:
        print(__doc__.strip())
        return 1
    pack, out = sys.argv[1], sys.argv[2]
    deep = int(sys.argv[3]) if len(sys.argv) > 3 else 2
    work = os.path.dirname(os.path.abspath(out))

    vals = openpyxl.load_workbook(pack, data_only=True)[SHEET]
    raw = openpyxl.load_workbook(pack)[SHEET]
    if str(raw['A5'].value).strip() != 'Line' or str(raw['O5'].value).strip() != 'YTD':
        print('STOPPED: %s does not start where expected' % SHEET)
        return 1

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    keep = []
    # row 1 is the sheet's own 'Profit & Loss' heading. The slide already has a
    # heading, and two headings on one slide is wrong, so it comes off.
    for r in range(2, LASTROW + 1):
        dim = raw.row_dimensions.get(r)
        if ((dim.outlineLevel or 0) if dim else 0) > deep:
            continue
        keep.append(r)
    for i, r in enumerate(keep, 1):
        for c in range(1, LASTCOL + 1):
            t, o = ws.cell(i, c), raw.cell(r, c)
            t.value = vals.cell(r, c).value
            t.font = copy.copy(o.font)
            t.fill = copy.copy(o.fill)
            t.border = copy.copy(o.border)
            t.alignment = copy.copy(o.alignment)
            t.number_format = o.number_format
        h = raw.row_dimensions[r].height if r in raw.row_dimensions else None
        ws.row_dimensions[i].height = h or 14.5
    for c in range(1, LASTCOL + 1):
        ws.column_dimensions[L(c)].width = raw.column_dimensions[L(c)].width or 12
    ws.sheet_view.showGridLines = False
    ws.print_area = 'A1:%s%d' % (L(LASTCOL), len(keep))
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.15
    ws.page_margins.top = ws.page_margins.bottom = 0.15
    tmp = os.path.join(work, 'plfy27.xlsx')
    wb.save(tmp)

    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', work, tmp],
                   check=True, capture_output=True)
    pdf = os.path.join(work, 'plfy27.pdf')
    subprocess.run(['pdftoppm', '-r', '300', '-png', '-f', '1', '-l', '1',
                    pdf, os.path.join(work, 'plfy27')], check=True, capture_output=True)
    page = os.path.join(work, 'plfy27-1.png')
    im = Image.open(page).convert('RGB')
    bg = Image.new('RGB', im.size, (255, 255, 255))
    bb = ImageChops.difference(im, bg).getbbox()
    pad = 18
    im.crop((max(0, bb[0] - pad), max(0, bb[1] - pad),
             min(im.size[0], bb[2] + pad), min(im.size[1], bb[3] + pad))).save(out)
    w, h = Image.open(out).size
    print('  %d of %d rows kept (outline level %d and above), %d x %d, aspect %.2f'
          % (len(keep), LASTROW, deep, w, h, w / h))
    print('  saved %s' % out)
    return 0


if __name__ == '__main__':
    sys.exit(main())
