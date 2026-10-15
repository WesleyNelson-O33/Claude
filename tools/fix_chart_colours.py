"""Put the chart furniture back on the CTS palette after a recalc.

LibreOffice writes its own black and greys into the chart XML for axis text,
tick marks and gridlines. They are not data colours, but they are not hers
either, so map them onto the nearest colour she gave.

    python3 fix_chart_colours.py "Workbook.xlsx"

Every other part of the zip is copied through byte for byte.
"""
import json, os, re, shutil, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
B = json.load(open(f'{HERE}/../pack/data/brand.json'))
ALLOWED = {v.upper() for k, v in B.items() if k != '_note'}
MAP = {'000000': B['charcoal'], '1A1A1A': B['charcoal'], '595959': B['charcoal'],
       '808080': B['grey'], 'B3B3B3': B['grey'], 'D9D9D9': B['grey'],
       'BFBFBF': B['grey'], 'F2F2F2': B['grey_lt'], 'CCCCCC': B['grey']}

def main(path):
    tmp = path + '.tmp'
    changed = 0
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(
            tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            data = zin.read(it.filename)
            if re.match(r'xl/charts/chart\d+\.xml$', it.filename):
                txt = data.decode('utf-8')

                def sub(m):
                    nonlocal changed
                    h = m.group(1).upper()
                    if h in ALLOWED:
                        return m.group(0)
                    changed += 1
                    return 'srgbClr val="%s"' % MAP.get(h, B['charcoal']).lower()

                txt = re.sub(r'srgbClr val="([0-9A-Fa-f]{6})"', sub, txt)
                data = txt.encode('utf-8')
            zout.writestr(it, data)
    shutil.move(tmp, path)
    print(f'chart colours moved onto the CTS palette: {changed}')
    return 0

if __name__ == '__main__':
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    sys.exit(main(sys.argv[1]))
