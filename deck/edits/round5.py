"""Round five.

1. Slide 8 picks up the rebuilt chart, which now carries the departments as
   stacked columns, and the commentary says how to read them.
2. Slide 3 gets a commentary panel. It was the only content slide without one.
   The summary paragraph was already there, unlabelled and sitting under a
   band of white, so the cards and the strip move up and the paragraph goes
   into a proper panel in the room that frees.
3. Every figure in every commentary panel is rounded to $xxk, as asked.
"""
import sys, os, re, copy
import openpyxl
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Emu

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from herdeck import title_of, fill_box

HERE = os.path.dirname(os.path.abspath(__file__))
SNIP = os.path.join(HERE, '..', 'snips')
PACK = '/home/user/Claude/pack/CTS Financial Controller Pack - v20.xlsx'
BODY = ('Text 5', 'Text 5b', 'Text 53')

# $1,234,567 -> $1,235k. An amount already in k is left alone, and so is
# anything under a thousand, because $500 would otherwise become $0k.
# the number has to end on a digit. [\d,]* is greedy and swallowed the
# comma after a figure, turning "$235,213, down" into "$235k down".
MONEY = re.compile(r'\$(\()?(\d(?:[\d,]*\d)?(?:\.\d+)?)(\))?(k\b)?')


def to_k(text):
    def sub(m):
        if m.group(4):
            return m.group(0)
        try:
            v = float(m.group(2).replace(',', ''))
        except ValueError:
            return m.group(0)
        if v < 1000:
            return m.group(0)
        return '$%s%sk%s' % (m.group(1) or '', format(round(v / 1000), ','),
                             m.group(3) or '')
    return MONEY.sub(sub, text)


def k(v):
    n = round(v / 1000)
    return ('$(%s)k' % format(abs(n), ',')) if n < 0 else ('$%sk' % format(n, ','))


def swap_images(pres):
    """Refresh every chart image in place, keeping her positions: the top left
    and the width stay, only the height follows the new aspect."""
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


def donor(pres):
    """Borrow the panel and the COMMENTARY heading off a slide that has them,
    so slide 3 gets the same shape, fill and type as every other panel."""
    for s in pres.slides:
        by = {sh.name: sh for sh in s.shapes}
        if 'commentary panel' in by and 'Text 4' in by:
            return by['commentary panel'], by['Text 4']
    return None, None


def add_panel_to_summary(pres):
    panel_src, head_src = donor(pres)
    if panel_src is None:
        print('  no donor panel found, slide 3 left alone')
        return
    for i, s in enumerate(pres.slides, 1):
        if not title_of(s).lower().startswith('performance summary'):
            continue
        by = {sh.name: sh for sh in s.shapes}
        if 'commentary panel' in by:
            print('  slide %d already has a panel' % i)
            return
        para = by.get('Text 53')
        if para is None:
            print('  slide %d has no summary paragraph' % i)
            return

        # close the band of white under the title. The cards and the strip
        # move up by different amounts because the gap between them stays.
        for sh in s.shapes:
            t = Emu(sh.top).inches
            if sh.name in ('cts logo', 'Text 0') or sh is para:
                continue
            if 1.55 <= t <= 3.40:
                sh.top = sh.top + Inches(-0.57)
            elif 3.50 <= t <= 4.15:
                sh.top = sh.top + Inches(-0.75)

        tree = s.shapes._spTree
        panel = copy.deepcopy(panel_src._element)
        head = copy.deepcopy(head_src._element)
        tree.append(panel)
        tree.append(head)
        # a cloned shape carries the donor's position, so set all of it
        new = {sh.name: sh for sh in s.shapes}
        p, h = new['commentary panel'], new['Text 4']
        p.left, p.top, p.width, p.height = (Inches(0.55), Inches(3.52),
                                            Inches(8.90), Inches(1.32))
        h.left, h.top, h.width, h.height = (Inches(0.75), Inches(3.60),
                                            Inches(3.00), Inches(0.25))
        para.left, para.top = Inches(0.75), Inches(3.90)
        para.width, para.height = Inches(8.50), Inches(1.05)
        # the panel is a fill, so it has to sit behind the writing
        tree.remove(panel)
        tree.insert(2, panel)
        print('  slide %d given a commentary panel' % i)
        return


def round_commentary(pres):
    n = 0
    for i, s in enumerate(pres.slides, 1):
        for sh in s.shapes:
            if not sh.has_text_frame or sh.name not in BODY:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    new = to_k(r.text)
                    if new != r.text:
                        r.text = new
                        n += 1
        if s.has_notes_slide:
            tf = s.notes_slide.notes_text_frame
            tf.text = to_k(tf.text)
    return n


def main(src, out):
    wb = openpyxl.load_workbook(PACK, data_only=True)
    sm, xh = wb['Summary'], wb['Xero History']
    ytd, bud = sm.cell(6, 5).value, sm.cell(6, 7).value
    short = bud - ytd
    fy25, fy26 = xh.cell(60, 2).value, xh.cell(60, 3).value

    pres = Presentation(src)
    print('  %d chart images refreshed' % swap_images(pres))

    for i, s in enumerate(pres.slides, 1):
        if not title_of(s).lower().startswith('year to date revenue'):
            continue
        by = {sh.name: sh for sh in s.shapes}
        fill_box(by['Text 5'], [
            'The columns are revenue by department. They stack to the FY27 line.',
            'Year to date revenue is %s against a budget of %s, %s or %.0f%% behind.'
            % (k(ytd), k(bud), k(short), short / bud * 100),
            'PwC was budgeted at $336k in August. That work will not now happen.',
        ])
        fill_box(by['Text 5b'], [
            'Bankwest was budgeted at $49k in September and has moved into Q2.',
            'Revenue is still ahead of both completed years, %s in FY25 and %s in FY26.'
            % (k(fy25), k(fy26)),
            'The dashed line is the full year plan. Nine months are still to come.',
        ])
        print('  slide %d commentary rewritten' % i)

    add_panel_to_summary(pres)
    print('  %d figures rounded to $xxk' % round_commentary(pres))
    pres.save(out)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
