"""Put the Q1 Dept Report on slide two, with the commentary and the notes.

Slide two was set up with its title and nothing else. The rest of the deck's
light pages come in two shapes: the commentary in a strip across the bottom,
or the commentary in a panel down the right. The report is a wide table, and
at the full 8.90 width it would stand five inches tall, which does not fit
under a title on a 5.62 inch slide. So this follows the right hand panel,
the way slide four carries its table, and the panel and its label are copied
off that slide so the look is hers and not mine.

Nothing on this slide is typed. Every figure in the commentary and in the
notes is read out of the controller pack when the slide is written, which is
the only way they stay right when the pack is re-pasted.

This is a quarterly slide. It talks about the quarter and never refers across
to the month, because the monthly slides may be hidden when it is presented.

    deck/edits/slide2_deptreport.py <deck.pptx> <snip.png> [out.pptx]
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
SHEET = 'Q1 Dept Report'
MODEL = 3                      # slide four, the page whose panel this copies
TARGET = 1                     # slide two

# the area the table gets, which is slide four's, and the panel beside it
AREA = dict(left=0.25, top=0.92, width=6.67, height=4.18)


def money(v):
    return ('-$%s' % format(abs(round(v)), ',')) if v < 0 else '$%s' % format(round(v), ',')


def amt(v):
    """The size of a figure, for a sentence that already says lost or loss."""
    return '$%s' % format(abs(round(v)), ',')


def pc(v):
    return '%.1f%%' % (v * 100)


def figures():
    ws = openpyxl.load_workbook(PACK, data_only=True)[SHEET]

    def g(ref):
        v = ws[ref].value
        if v is None or isinstance(v, str):
            raise SystemExit('STOPPED: %s!%s is %r, the sheet has moved' % (SHEET, ref, v))
        return v
    return dict(
        rev27=g('L4'), rev26=g('M4'), gp27=g('L11'), gp26=g('M11'),
        gpc27=g('L12'), gpc26=g('M12'),
        int27=g('H4'), int26=g('I4'), intgpc=g('H12'),
        sup=g('B27'), sup26=g('C29'), prd=g('D27'), prd26=g('E29'),
        con=g('H27'), con26=g('I29'), cts=g('L27'), cts26=g('M29'),
        evgpc27=g('D12'), evgpc26=g('E12'),
        pe_sup=g('B37'), pe_prd=g('D37'), pe_con=g('H37'))


def bullets(f):
    return [
        'Revenue was %s for the quarter against %s, up %s.'
        % (money(f['rev27']), money(f['rev26']),
           pc(f['rev27'] / f['rev26'] - 1)),
        'Gross profit was %s against %s, but the margin fell from %s to %s.'
        % (money(f['gp27']), money(f['gp26']), pc(f['gpc26']), pc(f['gpc27'])),
        'Integration drove it, %s against %s, at a %s margin.'
        % (money(f['int27']), money(f['int26']), pc(f['intgpc'])),
        'Support is the only department in profit after overheads, %s.'
        % money(f['sup']),
        'Production lost %s and Consulting %s.' % (amt(f['prd']), amt(f['con'])),
        'The group lost %s, against a loss of %s last year.'
        % (amt(f['cts']), amt(f['cts26'])),
        'Profit per FTE: Support %s, Production %s, Consulting %s.'
        % (money(f['pe_sup']), money(f['pe_prd']), money(f['pe_con'])),
    ]


def notes(f):
    return '\n'.join([
        'Revenue for the first quarter was %s against %s last year, %s more, up %s.'
        % (money(f['rev27']), money(f['rev26']),
           money(f['rev27'] - f['rev26']), pc(f['rev27'] / f['rev26'] - 1)),
        'Gross profit was %s against %s. So revenue grew a long way faster than '
        'gross profit, and the margin came down from %s to %s.'
        % (money(f['gp27']), money(f['gp26']), pc(f['gpc26']), pc(f['gpc27'])),
        '',
        'Support had the steady quarter. It is the only department that makes a '
        'profit once overheads are on it, %s for the quarter against %s last year.'
        % (money(f['sup']), money(f['sup26'])),
        '',
        'Production improved its margin, Events from %s to %s, but it still '
        'finishes at a loss of %s after overheads, against %s last year.'
        % (pc(f['evgpc26']), pc(f['evgpc27']), amt(f['prd']), amt(f['prd26'])),
        '',
        'Consulting is where the growth is. Integration revenue was %s against %s. '
        'The margin on it is %s, so the volume is large and the gross profit on it '
        'is small. After overheads Consulting loses %s.'
        % (money(f['int27']), money(f['int26']), pc(f['intgpc']), amt(f['con'])),
        'That one department is the reason the group margin fell.',
        '',
        'The group lost %s for the quarter, against a loss of %s last year. So we '
        'are close to breakeven and a long way better than last year.'
        % (amt(f['cts']), amt(f['cts26'])),
        'Profit per full time equivalent was %s in Support, %s in Production and '
        '%s in Consulting.'
        % (money(f['pe_sup']), money(f['pe_prd']), money(f['pe_con'])),
        '',
        'If you are asked: last year had no departmental overhead split, so the '
        'FY26 half of this table is the published FY26 report, pasted in. The '
        'employee counts are full time equivalents off labour hours, the same '
        'basis the old report used, so the two years compare.',
    ])


def main():
    if len(sys.argv) < 3:
        print(__doc__.strip())
        return 1
    src, snip = sys.argv[1], sys.argv[2]
    out = sys.argv[3] if len(sys.argv) > 3 else src
    f = figures()

    pres = Presentation(src)
    if len(pres.slides) <= max(MODEL, TARGET):
        print('STOPPED: the deck has only %d slides' % len(pres.slides))
        return 1
    model, slide = pres.slides[MODEL], pres.slides[TARGET]
    by = {sh.name: sh for sh in model.shapes}
    for want in ('commentary panel', 'Text 4', 'Text 5'):
        if want not in by:
            print('STOPPED: slide %d has no %r to copy' % (MODEL + 1, want))
            return 1

    tree = slide.shapes._spTree
    for name in ('commentary panel', 'Text 4', 'Text 5'):
        for old in [sh for sh in slide.shapes if sh.name == name]:
            tree.remove(old._element)
        tree.append(copy.deepcopy(by[name]._element))
    now = {sh.name: sh for sh in slide.shapes}

    # the table, centred in what is left beside the panel
    im = Image.open(snip)
    aspect = im.width / im.height
    w = min(AREA['width'], AREA['height'] * aspect)
    h = w / aspect
    pic = slide.shapes.add_picture(
        snip, Inches(AREA['left'] + (AREA['width'] - w) / 2),
        Inches(AREA['top'] + (AREA['height'] - h) / 2),
        width=Inches(w), height=Inches(h))
    pic.name = 'q1 dept report'
    tree.remove(pic._element)
    tree.insert(2, pic._element)

    if not any(sh.name == 'cts logo' for sh in slide.shapes):
        logo = slide.shapes.add_picture(LOGO, Inches(8.77), Inches(0.26),
                                        width=Inches(0.68), height=Inches(0.38))
        logo.name = 'cts logo'

    box = now['Text 5']
    write(box, bullets(f))
    for para in box.text_frame.paragraphs:
        for r in para.runs:
            r.font.size = Pt(SIZE)

    tf = slide.notes_slide.notes_text_frame
    if tf is None:
        # a notes page made from scratch comes out without a body box, so the
        # one off slide four, which already carries speaking notes, is copied in
        donor = model.notes_slide.notes_text_frame
        if donor is None:
            print('STOPPED: slide %d has no notes box to copy either' % (MODEL + 1))
            return 1
        slide.notes_slide.shapes._spTree.append(copy.deepcopy(donor._parent._element))
        tf = slide.notes_slide.notes_text_frame
        if tf is None:
            print('STOPPED: the notes box did not take')
            return 1
    tf.text = notes(f)
    pres.save(out)
    print('  slide %d: table %.2f x %.2f in, %d commentary lines, notes written'
          % (TARGET + 1, w, h, len(bullets(f))))
    print('  saved %s' % out)
    return 0


if __name__ == '__main__':
    sys.exit(main())
