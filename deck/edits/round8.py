"""Round eight: slide 8's columns are reporting groups now, so say so.

Also refreshes every chart image, because the pack's combined charts had been
coming out of LibreOffice with their lines stacked and have been repaired.
"""
import sys, os
import openpyxl
from PIL import Image
from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from herdeck import title_of, fill_box

HERE = os.path.dirname(os.path.abspath(__file__))
SNIP = os.path.join(HERE, '..', 'snips')
PACK = '/home/user/Claude/pack/CTS Financial Controller Pack - v20.xlsx'


def k(v):
    n = round(v / 1000)
    return ('$(%s)k' % format(abs(n), ',')) if n < 0 else ('$%sk' % format(n, ','))


def swap_images(pres):
    n = 0
    for slide in pres.slides:
        for sh in list(slide.shapes):
            if 'PICTURE' not in str(sh.shape_type) or not sh.name.startswith('snip '):
                continue
            png = os.path.join(SNIP, sh.name[5:] + '.png')
            if not os.path.exists(png):
                continue
            left, top, width, name = sh.left, sh.top, sh.width, sh.name
            tree = slide.shapes._spTree
            tree.remove(sh._element)
            im = Image.open(png)
            pic = slide.shapes.add_picture(png, left, top, width=width,
                                           height=int(width * im.height / im.width))
            pic.name = name
            tree.remove(pic._element)
            tree.insert(2, pic._element)
            n += 1
    return n


def main(src, out):
    wb = openpyxl.load_workbook(PACK, data_only=True)
    sm, xh = wb['Summary'], wb['Xero History']
    ytd, bud = sm.cell(6, 5).value, sm.cell(6, 7).value
    short = bud - ytd
    fy25, fy26 = xh.cell(60, 2).value, xh.cell(60, 3).value
    groups = [(xh.cell(257, c).value, xh.cell(260, c).value) for c in (2, 3, 4)]

    pres = Presentation(src)
    print('  %d chart images refreshed' % swap_images(pres))
    for i, s in enumerate(pres.slides, 1):
        if not title_of(s).lower().startswith('year to date revenue'):
            continue
        by = {sh.name: sh for sh in s.shapes}
        fill_box(by['Text 5'], [
            'The columns are revenue by reporting group. They stack to the FY27 line.',
            '%s.' % ', '.join('%s %s' % (n, k(v)) for n, v in groups),
            'Year to date revenue is %s against a budget of %s, %s or %.0f%% behind.'
            % (k(ytd), k(bud), k(short), short / bud * 100),
        ])
        fill_box(by['Text 5b'], [
            'PwC was budgeted at $336k in August. That work will not now happen.',
            'Bankwest was budgeted at $49k in September and has moved into Q2.',
            'Revenue is still ahead of both completed years, %s in FY25 and %s in FY26.'
            % (k(fy25), k(fy26)),
        ])
        print('  slide %d commentary rewritten' % i)
        print('     %s' % ', '.join('%s %s' % (n, k(v)) for n, v in groups))
        break
    pres.save(out)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
