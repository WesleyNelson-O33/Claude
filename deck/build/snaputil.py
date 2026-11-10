"""Snip the Jul-Sept sheet of the utilisation report.

It is not part of the controller pack, so it gets its own small snipper. Same
settings as snip.py: landscape, fit to width, no header or footer, trimmed.
"""
import os, subprocess, sys
# -I keeps this script's own directory off sys.path, so the import of trim
# from snip.py below fails without this line.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.worksheet.properties import PageSetupProperties

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(HERE, 'snipwork')
OUT = os.path.join(HERE, '..', 'snips')
SOFFICE = os.path.join(HERE, 'soffice_shim.py') if os.path.exists(
    os.path.join(HERE, 'soffice_shim.py')) else None


def main(src, sheet, rng, snip_id, hide_cols=''):
    os.makedirs(WORK, exist_ok=True)
    wb = openpyxl.load_workbook(src)
    for ws in wb.worksheets:
        ws.sheet_state = 'visible' if ws.title == sheet else 'hidden'
    wb.active = wb.index(wb[sheet])
    ws = wb[sheet]
    for col in hide_cols:
        ws.column_dimensions[col].hidden = True
    ws.print_area = rng
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.gridLines = False
    for side in ('left', 'right', 'top', 'bottom'):
        setattr(ws.page_margins, side, 0.15)
    ws.page_margins.header = ws.page_margins.footer = 0
    for hf in (ws.oddHeader, ws.oddFooter, ws.evenHeader, ws.evenFooter,
               ws.firstHeader, ws.firstFooter):
        for part in (hf.left, hf.center, hf.right):
            part.text = None
    ws.HeaderFooter.differentFirst = False
    ws.HeaderFooter.differentOddEven = False
    x = os.path.join(WORK, snip_id + '.xlsx')
    wb.save(x)
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', WORK, x],
                   check=True, capture_output=True)
    pdf = os.path.join(WORK, snip_id + '.pdf')
    png = os.path.join(OUT, snip_id + '.png')
    subprocess.run(['pdftoppm', '-f', '1', '-l', '1', '-r', '200', '-png',
                    '-singlefile', pdf, png[:-4]], check=True, capture_output=True)
    from snip import trim
    trim(png)
    from PIL import Image
    im = Image.open(png)
    print('  %-12s %-18s %-12s %dx%d' % (snip_id, sheet, rng, im.width, im.height))


if __name__ == '__main__':
    main(*sys.argv[1:])
