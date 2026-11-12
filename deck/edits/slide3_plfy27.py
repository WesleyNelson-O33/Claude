"""Add the FY27 P&L by month after the Q1 department page.

This is a MONTHLY slide. It carries the twelve monthly columns of the P&L
plus the FY total and the year to date, which is columns A to O of the
'P&L FY27' tab. It says nothing about the quarter, so the quarterly pages
either side can be hidden without it reading oddly.

It is laid out the same way as the page before it, the table on the left
and the commentary down the right, because two table pages in a row that
are built differently look like a mistake. The template is that page, so
every colour, font and box position is hers.

Nothing is typed. Every figure in the commentary and the notes is read out
of the pack when the slide is written.

    deck/edits/slide3_plfy27.py <deck.pptx> <snip.png> [out.pptx]
"""
import copy
import os
import sys

import openpyxl
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from addslide import LOGO, SIZE, write

PACK = '/home/user/Claude/pack/CTS Financial Controller Pack - v20.xlsx'
SHEET = 'P&L FY27'
MODEL = 1                      # slide two, the page this copies
AT = 2                         # it becomes slide three
AREA = dict(left=0.25, top=0.92, width=6.55, height=4.18)  # 0.12 clear of the panel
EYEBROW, TITLE = 'FINANCIAL PERFORMANCE', 'FY27 P&L BY MONTH'
MONTHS = ('B', 'C', 'D')       # July, August, September, the actual months


def money(v):
    return ('-$%s' % format(abs(round(v)), ',')) if v < 0 else '$%s' % format(round(v), ',')


def amt(v):
    return '$%s' % format(abs(round(v)), ',')


def pc(v):
    return '%.1f%%' % (v * 100)


def figures():
    ws = openpyxl.load_workbook(PACK, data_only=True)[SHEET]
    if str(ws['A5'].value).strip() != 'Line' or str(ws['O5'].value).strip() != 'YTD':
        raise SystemExit('STOPPED: %s does not start where expected' % SHEET)

    def g(ref):
        v = ws[ref].value
        if v is None or isinstance(v, str):
            raise SystemExit('STOPPED: %s!%s is %r, the sheet has moved' % (SHEET, ref, v))
        return v
    rows = dict(inc=6, oh=164, gp=173, gpc=174, net=181, netc=182)
    f = {}
    for key, r in rows.items():
        f[key + '_ytd'] = g('O%d' % r)
        f[key + '_fy'] = g('N%d' % r)
        f[key + '_m'] = [g('%s%d' % (c, r)) for c in MONTHS]
    f['names'] = [ws['%s5' % c].value.strftime('%B') for c in MONTHS]
    f['month'] = ws['B3'].value.strftime('%B %Y')
    return f


def bullets(f):
    best = max(range(3), key=lambda i: f['net_m'][i])
    return [
        'The year to date is %s of revenue and a net loss of %s.'
        % (money(f['inc_ytd']), amt(f['net_ytd'])),
        'By month, revenue was %s, %s and %s.'
        % tuple(money(v) for v in f['inc_m']),
        'Only %s made a profit, %s. %s lost %s and %s lost %s.'
        % (f['names'][best], money(f['net_m'][best]),
           f['names'][0], amt(f['net_m'][0]), f['names'][2], amt(f['net_m'][2])),
        'Gross margin by month was %s, %s and %s.'
        % tuple(pc(v) for v in f['gpc_m']),
        'Overheads barely move, %s, %s and %s a month.'
        % tuple(amt(v) for v in f['oh_m']),
        'The forecast carries %s of revenue for the full year and %s of net profit.'
        % (money(f['inc_fy']), money(f['net_fy'])),
        'Everything after %s is forecast, not actual.' % f['month'],
    ]


def notes(f):
    best = max(range(3), key=lambda i: f['net_m'][i])
    worst = min(range(3), key=lambda i: f['net_m'][i])
    return '\n'.join([
        'This is the P&L by month for the full year. The first three columns are '
        'actual. Everything from the fourth column on is forecast, so do not let '
        'the room read those as results.',
        '',
        'Revenue year to date is %s. Net result is a loss of %s.'
        % (money(f['inc_ytd']), amt(f['net_ytd'])),
        '',
        'The three actual months were %s, %s and %s. Only %s made a profit, %s. '
        '%s lost %s and %s lost %s.'
        % (money(f['inc_m'][0]), money(f['inc_m'][1]), money(f['inc_m'][2]),
           f['names'][best], money(f['net_m'][best]),
           f['names'][0], amt(f['net_m'][0]), f['names'][2], amt(f['net_m'][2])),
        '',
        'The reason is margin, not cost. Gross margin ran %s, %s and %s. %s is the '
        'weakest of the three. Overheads were %s, %s and %s, so they are steady and '
        'they are not what moved the result.'
        % (pc(f['gpc_m'][0]), pc(f['gpc_m'][1]), pc(f['gpc_m'][2]), f['names'][worst],
           amt(f['oh_m'][0]), amt(f['oh_m'][1]), amt(f['oh_m'][2])),
        '',
        'On the full year the forecast shows %s of revenue and %s of net profit at a '
        '%s gross margin. That is the number to question, because it needs the back '
        'half of the year to run well ahead of what the first three months did.'
        % (money(f['inc_fy']), money(f['net_fy']), pc(f['gpc_fy'])),
        '',
        'If you are asked why the account detail is not shown: it is there in the '
        'pack, this is the same sheet with the groups collapsed so it can be read '
        'on a slide.',
    ])


def main():
    if len(sys.argv) < 3:
        print(__doc__.strip())
        return 1
    src, snip = sys.argv[1], sys.argv[2]
    out = sys.argv[3] if len(sys.argv) > 3 else src
    f = figures()

    pres = Presentation(src)
    model = pres.slides[MODEL]
    new = pres.slides.add_slide(model.slide_layout)
    for sh in list(new.shapes):
        sh._element.getparent().remove(sh._element)
    for sh in model.shapes:
        new.shapes._spTree.append(copy.deepcopy(sh._element))

    tree = new.shapes._spTree
    for sh in list(new.shapes):                 # pictures carry a rel id, not the image
        if 'PICTURE' in str(sh.shape_type):
            tree.remove(sh._element)
    by = {sh.name: sh for sh in new.shapes}
    for want in ('Text 0', 'Text 4', 'Text 5', 'commentary panel'):
        if want not in by:
            print('STOPPED: slide %d has no %r to copy' % (MODEL + 1, want))
            return 1

    runs = by['Text 0'].text_frame.paragraphs[0].runs
    if len(runs) < 3:
        print('STOPPED: the title on slide %d is not an eyebrow, a bar and a title'
              % (MODEL + 1))
        return 1
    runs[0].text, runs[2].text = EYEBROW, TITLE

    im = Image.open(snip)
    aspect = im.width / im.height
    w = min(AREA['width'], AREA['height'] * aspect)
    h = w / aspect
    pic = new.shapes.add_picture(
        snip, Inches(AREA['left'] + (AREA['width'] - w) / 2),
        Inches(AREA['top'] + (AREA['height'] - h) / 2),
        width=Inches(w), height=Inches(h))
    pic.name = 'p&l fy27'
    tree.remove(pic._element)
    tree.insert(2, pic._element)
    logo = new.shapes.add_picture(LOGO, Inches(8.77), Inches(0.26),
                                  width=Inches(0.68), height=Inches(0.38))
    logo.name = 'cts logo'

    box = by['Text 5']
    write(box, bullets(f))
    for para in box.text_frame.paragraphs:
        for r in para.runs:
            r.font.size = Pt(SIZE)

    ns = new.notes_slide
    if ns.notes_text_frame is None:
        donor = model.notes_slide.notes_text_frame
        ns.shapes._spTree.append(copy.deepcopy(donor._parent._element))
    new.notes_slide.notes_text_frame.text = notes(f)

    ids = pres.slides._sldIdLst
    el = ids[-1]
    ids.remove(el)
    ids.insert(AT, el)

    pres.save(out)
    print('  slide %d added: table %.2f x %.2f in, %d commentary lines'
          % (AT + 1, w, h, len(bullets(f))))
    print('  saved %s' % out)
    return 0


if __name__ == '__main__':
    sys.exit(main())
