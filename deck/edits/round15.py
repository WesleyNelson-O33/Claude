"""Add a line to the utilisation slide's presenter notes.

Danica asked for it to be said that Blake's leave sits inside Production's
leave total, so nobody reads Production's hours without knowing why they are
short. The figures are read from the pack's Utilisation Q1 sheet.
"""
import sys, os
import openpyxl
from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from herdeck import title_of

PACK = '/home/user/Claude/pack/CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Utilisation Q1'
TARGET = 'utilisation per employee'
WHO = 'Blake Crisford'
GROUP_TOTAL = 'Production Total'


def main(src, out):
    ws = openpyxl.load_workbook(PACK, data_only=True)[SHEET]
    rows = {}
    for r in range(3, ws.max_row + 1):
        name = ws.cell(r, 2).value or ws.cell(r, 1).value
        if name:
            rows.setdefault(str(name).strip(), r)
    for key in (WHO, GROUP_TOTAL):
        if key not in rows:
            print('  STOPPED: no row for %r on %s' % (key, SHEET))
            return
    his = ws.cell(rows[WHO], 5).value              # leave hours
    grp = ws.cell(rows[GROUP_TOTAL], 5).value
    line = ('%s’s leave is included in Production’s leave time. He took %s '
            'of Production’s %s leave hours for the quarter, which is %.0f%% of it. '
            'Production’s hours look short partly because of that.'
            % (WHO, format(round(his), ','), format(round(grp), ','), his / grp * 100))

    pres = Presentation(src)
    for i, s in enumerate(pres.slides, 1):
        if not title_of(s).strip().lower().startswith(TARGET):
            continue
        tf = s.notes_slide.notes_text_frame
        have = tf.text or ''
        if WHO in have:
            print('  slide %d already says it' % i)
        else:
            tf.text = (have.rstrip() + '\n' + line) if have.strip() else line
            print('  slide %d notes extended' % i)
        print('  %s' % line)
        break
    else:
        print('  utilisation slide not found')
        return
    pres.save(out)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
