"""Rebuild the profit per employee commentary from the table.

The year labels moved by one when Danica relabelled the block FY23 to FY27.
The commentary named years, so it started pointing at the wrong rows: PRD's
peak was called FY23 when it is now FY24. Writing it from the figures rather
than editing the words means it cannot drift again.
"""
import sys, os
import openpyxl
from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from herdeck import title_of, fill_box

PACK = '/home/user/Claude/pack/CTS Financial Controller Pack - v20.xlsx'
SHEET, FIRST, LAST = 'Xero History', 153, 157
COLS = {'ONS': 2, 'PRD': 3, 'CONS': 4, 'CTS': 5}
TARGET = 'profit per employee'


def k(v):
    n = round(v / 1000)
    return ('$(%sk)' % format(abs(n), ',')) if n < 0 else ('$%sk' % format(n, ','))


def main(src, out):
    ws = openpyxl.load_workbook(PACK, data_only=True)[SHEET]
    yrs = [ws.cell(r, 1).value for r in range(FIRST, LAST + 1)]
    col = {g: [ws.cell(r, c).value for r in range(FIRST, LAST + 1)]
           for g, c in COLS.items()}
    now, prev = -1, -2

    peak_i = max(range(len(yrs)), key=lambda i: col['PRD'][i])
    cons = col['CONS']
    worse = [i for i in range(len(yrs) - 1) if cons[i] < cons[now]]
    worst_since = yrs[max(worse)] if worse else None
    times = cons[now] / cons[prev] if cons[prev] else 0
    negatives = sum(1 for v in cons if v < 0)

    left = [
        'ONS earned %s a head in %s, down from %s the year before.'
        % (k(col['ONS'][now]), yrs[now], k(col['ONS'][prev])),
        'PRD recovered to %s from %s, still a long way under the %s it reached in %s.'
        % (k(col['PRD'][now]), k(col['PRD'][prev]), k(col['PRD'][peak_i]), yrs[peak_i]),
        'CTS overall earned %s a head, down from %s.'
        % (k(col['CTS'][now]), k(col['CTS'][prev])),
    ]
    right = [
        'CONS lost %s a head%s.'
        % (k(abs(cons[now])),
           ', its worst since %s' % worst_since if worst_since else ''),
        'That is %.1f times the loss it made in %s.' % (times, yrs[prev]),
        'CONS has not earned a positive figure in any of the %d years on this chart.'
        % negatives if negatives == len(yrs) else
        'CONS was negative in %d of the %d years on this chart.' % (negatives, len(yrs)),
    ]

    pres = Presentation(src)
    for i, s in enumerate(pres.slides, 1):
        if not title_of(s).strip().lower().startswith(TARGET):
            continue
        by = {sh.name: sh for sh in s.shapes}
        fill_box(by['Text 5'], left)
        fill_box(by['Text 5b'], right)
        try:
            s.notes_slide.notes_text_frame.text = '\n'.join(left + right)
        except AttributeError:
            pass
        print('  slide %d commentary rebuilt from %s rows %d-%d'
              % (i, SHEET, FIRST, LAST))
        for t in left + right:
            print('    %s' % t)
        break
    pres.save(out)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
