"""Undo LibreOffice's habit of stacking the lines on a combined chart.

A chart that has columns and lines in the same plot area is written correctly
by openpyxl: the columns stacked, the lines standard. Running the workbook
through LibreOffice to fill in the formula cache rewrites the line plot as
stacked too, and the lines then draw on top of one another. On the year to
date chart that put FY26 at 6.8 million by December instead of 3.1 million.

This repairs it in the file, after the recalculation. A plot area holding both
columns and lines never wants stacked lines, so any it finds are put back to
standard. Nothing else in the workbook is touched.

Run it after tools/recalc.py, before taking any snips.
"""
import os, re, shutil, sys, tempfile, zipfile

CHART = re.compile(r'xl/charts/chart\d+\.xml$')
LINE = re.compile(r'(<(\w*:?)lineChart>.*?</\2lineChart>)', re.S)
GROUPING = re.compile(r'(<(\w*:?)grouping val=")stacked(")')


def fix(path):
    path = os.path.abspath(path)
    with zipfile.ZipFile(path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    changed = []
    out = []
    for info, data in items:
        if CHART.search(info.filename):
            x = data.decode('utf-8')
            if '<' in x and 'barChart' in x and 'lineChart' in x:
                def one(m):
                    body, n = GROUPING.subn(r'\1standard\3', m.group(1), count=1)
                    if n:
                        changed.append(info.filename)
                    return body
                new = LINE.sub(one, x)
                if new != x:
                    data = new.encode('utf-8')
        out.append((info, data))
    if not changed:
        print('no stacked lines found in %s' % os.path.basename(path))
        return 0
    tmp = tempfile.mktemp(suffix='.xlsx')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        for info, data in out:
            z.writestr(info, data)
    shutil.move(tmp, path)
    for c in changed:
        print('  %s: lines put back to standard' % c)
    return len(changed)


if __name__ == '__main__':
    for p in sys.argv[1:]:
        fix(p)
