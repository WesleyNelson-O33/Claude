"""Round ten: the customer slide becomes a three year comparison.

Danica asked for both views, so the chart takes the existing slide and the
same figures go on a new slide straight after it as a table. Both are
quarterly slides, so hiding one leaves the other reading on its own.

Every figure in the commentary is read from the pack's Top 10 Q1 3yr sheet.
"""
import sys, os, copy
import openpyxl
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Emu

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from herdeck import title_of, set_title, fill_box

HERE = os.path.dirname(os.path.abspath(__file__))
SNIP = os.path.join(HERE, '..', 'snips')
PACK = '/home/user/Claude/pack/CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Top 10 Q1 3yr'
FIRST, LAST = 3, 12                      # the ten rows
TOTAL = 13


def k(v):
    return '$%sk' % format(round(v / 1000), ',')


QUARTER = {'Q1 2024': 'Q1 FY25', 'Q1 2025': 'Q1 FY26', 'Q1 2026': 'Q1 FY27'}


def swap_images(pres):
    """Refresh every chart image in place, keeping her positions. The year
    labels changed in the pack, so every picture that carries one is stale."""
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


def relabel(pres):
    """The deck's own words used the calendar year for a quarter too."""
    n = 0
    for slide in pres.slides:
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            for para in sh.text_frame.paragraphs:
                for r in para.runs:
                    new = r.text
                    for was, now in QUARTER.items():
                        new = new.replace(was, now)
                    if new != r.text:
                        r.text = new
                        n += 1
    return n


def place(slide, name, png, top_lim, bot_lim):
    """Fit the picture to the room between the title and the panel, centred
    both ways, so no band of white is left on either side of it."""
    pic = next((s for s in slide.shapes if s.name == name), None)
    if pic is None:
        return None
    left = Emu(pic.left).inches
    room_w = 10.0 - 2 * left
    im = Image.open(png)
    w = room_w
    h = w * im.height / im.width
    if h > bot_lim - top_lim:
        h = bot_lim - top_lim
        w = h * im.width / im.height
    tree = slide.shapes._spTree
    tree.remove(pic._element)
    new = slide.shapes.add_picture(
        png, Inches(left + (room_w - w) / 2),
        Inches(top_lim + (bot_lim - top_lim - h) / 2),
        width=Inches(w), height=Inches(h))
    new.name = name
    tree.remove(new._element)
    tree.insert(2, new._element)
    return w, h


def clone_after(pres, src, index):
    new = pres.slides.add_slide(src.slide_layout)
    for sh in list(new.shapes):
        sh._element.getparent().remove(sh._element)
    for sh in src.shapes:
        new.shapes._spTree.append(copy.deepcopy(sh._element))
    lst = pres.slides._sldIdLst
    entry = list(lst)[-1]
    lst.remove(entry)
    lst.insert(index, entry)
    return new


def main(src, out):
    ws = openpyxl.load_workbook(PACK, data_only=True)[SHEET]
    names = {ws.cell(r, 1).value: r for r in range(FIRST, LAST + 1)}
    g = lambda n, c: ws.cell(names[n], c).value
    tot = [ws.cell(TOTAL, c).value for c in (2, 3, 4, 5)]
    new_names = [ws.cell(r, 1).value for r in range(FIRST, LAST + 1)
                 if ws.cell(r, 6).value == 'new this year']
    new_sum = sum(ws.cell(names[n], 4).value for n in new_names)

    pres = Presentation(src)
    print('  %d chart images refreshed' % swap_images(pres))
    print('  %d runs relabelled to the financial year' % relabel(pres))
    target, idx = None, None
    for i, s in enumerate(pres.slides):
        if title_of(s).lower().startswith('top ten customers'):
            target, idx = s, i
            break
    if target is None:
        print('  top ten slide not found')
        return

    def tighten(slide):
        by = {sh.name: sh for sh in slide.shapes}
        p, h = by.get('commentary panel'), by.get('Text 4')
        if p is None or h is None:
            return 4.0
        floor = Emu(p.top).inches + Emu(p.height).inches
        p.top, p.height = Inches(floor - 1.00), Inches(1.00)
        h.top, h.height = Inches(floor - 0.94), Inches(0.22)
        for nm in ('Text 5', 'Text 5b'):
            b = by.get(nm)
            if b is not None:
                b.top, b.height = Inches(floor - 0.70), Inches(0.64)
        return floor - 1.00

    top_lim = 0.94
    bot_lim = tighten(target) - 0.12

    # the chart keeps the slide
    set_title(target, 'Top Ten Customers, Three Year Comparison')
    got = place(target, 'snip top10',
                os.path.join(SNIP, 'top10_3yr.png'), top_lim, bot_lim)
    by = {sh.name: sh for sh in target.shapes}
    fill_box(by['Text 5'], [
        'The three bars are Q1 FY25, Q1 FY26 and Q1 FY27 for this quarter’s top ten.',
        'Deloitte has grown in all three, %s then %s then %s.'
        % (k(g('Deloitte Services', 2)), k(g('Deloitte Services', 3)),
           k(g('Deloitte Services', 4))),
        'CBA Place South has fallen in all three, %s down to %s.'
        % (k(g('CBA Place South', 2)), k(g('CBA Place South', 4))),
    ])
    fill_box(by['Text 5b'], [
        'ASX Operations has also fallen in all three, %s down to %s.'
        % (k(g('ASX Operations', 2)), k(g('ASX Operations', 4))),
        'Ventia and PwC are up in all three.',
        '%s of the ten were not customers two years ago: %s, %s between them.'
        % (len(new_names), ', '.join(new_names), k(new_sum)),
    ])
    print('  slide %d is the chart, picture %.2f x %.2f' % (idx + 1, got[0], got[1]))

    # the table gets its own slide straight after
    tbl = clone_after(pres, target, idx + 1)
    set_title(tbl, 'Top Ten Customers, The Figures')
    tighten(tbl)
    place(tbl, 'snip top10', os.path.join(SNIP, 'top10_3yr_tbl.png'),
          top_lim, bot_lim)
    by = {sh.name: sh for sh in tbl.shapes}
    fill_box(by['Text 5'], [
        'These ten are this quarter’s top ten. The earlier columns are what those '
        'same ten were worth then, not the top ten of those years.',
        'Together they were %s in Q1 FY25, %s in Q1 FY26 and %s this quarter.'
        % (k(tot[0]), k(tot[1]), k(tot[2])),
    ])
    fill_box(by['Text 5b'], [
        'That is %s more than last year, of which %s is the three new names.'
        % (k(tot[3]), k(new_sum)),
        'Q1 FY25 is not in the pack as a report. It was summed from the Xero '
        'invoices for that quarter, net of GST.',
    ])
    for s in (target, tbl):
        notes = []
        for nm in ('Text 5', 'Text 5b'):
            box = {x.name: x for x in s.shapes}.get(nm)
            if box is not None:
                notes += [''.join(r.text for r in p.runs)
                          for p in box.text_frame.paragraphs]
        try:
            s.notes_slide.notes_text_frame.text = '\n'.join(
                t for t in notes if t.strip())
        except AttributeError:
            pass
    print('  slide %d is the table' % (idx + 2))
    pres.save(out)
    print('  %d slides out' % len(pres.slides))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
