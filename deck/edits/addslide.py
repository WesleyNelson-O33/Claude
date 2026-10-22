"""Add a chart slide built from page 6, which already carries the layout the
rest of the deck uses: picture under the title, commentary in a strip below.

Cloning a slide copies each shape's relationship id but not the image behind
it, so every picture on the new slide is placed from file rather than copied.
"""
import copy, sys
from PIL import Image
from pptx import Presentation
from pptx.text.text import _Paragraph
from pptx.util import Inches, Pt

LOGO = '/home/user/Claude/deck/assets/cts-logo.png'
TOP, FLOOR, LEFT = 1.22, 5.06, 0.55
SIZE, LEAD, GAP = 10.0, 12.5, 4.0

def lines(bullets, w_in):
    cpl = max(12, int(32 * (10 / SIZE) * (w_in / 2.20)))
    return sum(max(1, -(-len(t) // cpl)) for t in bullets)

def need(bullets, w_in):
    return (lines(bullets, w_in) * LEAD + (len(bullets) - 1) * GAP) / 72 / 0.94

def balance(bullets, w_in):
    best, where = None, 1
    for k in range(1, len(bullets)):
        worst = max(lines(bullets[:k], w_in), lines(bullets[k:], w_in))
        if best is None or worst < best:
            best, where = worst, k
    return bullets[:where], bullets[where:]

def write(box, items):
    tf = box.text_frame
    first = tf.paragraphs[0]
    for para in list(tf.paragraphs)[1:]:
        para._p.getparent().remove(para._p)
    for r in list(first.runs)[1:]:
        r._r.getparent().remove(r._r)
    first.runs[0].text = items[0]
    for extra in items[1:]:
        np = copy.deepcopy(first._p)
        first._p.getparent().append(np)
        _Paragraph(np, tf).runs[0].text = extra

def add(pres, template_index, at, eyebrow, title, snip, bullets, notes, name):
    src = pres.slides[template_index]
    new = pres.slides.add_slide(src.slide_layout)
    for sh in list(new.shapes):
        sh._element.getparent().remove(sh._element)
    for sh in src.shapes:
        new.shapes._spTree.append(copy.deepcopy(sh._element))

    by = {sh.name: sh for sh in new.shapes}
    tree = new.shapes._spTree
    by['Text 0'].text_frame.paragraphs[0].runs[0].text = eyebrow
    by['Text 1'].text_frame.paragraphs[0].runs[0].text = title

    for sh in list(new.shapes):
        if 'PICTURE' not in str(sh.shape_type):
            continue
        geom = (sh.left, sh.top, sh.width, sh.height)
        is_logo = sh.name == 'cts logo'
        tree.remove(sh._element)
        if is_logo:
            mark = new.shapes.add_picture(LOGO, geom[0], geom[1],
                                          width=geom[2], height=geom[3])
            mark.name = 'cts logo'

    cols = list(balance(bullets, 4.05))
    ph_panel = max(need(c, 4.05) for c in cols if c) + 0.30
    strip_top = FLOOR - ph_panel

    im = Image.open(snip)
    aspect = im.width / im.height
    pic_h = min(strip_top - 0.12 - TOP, 8.90 / aspect)
    pic_w = pic_h * aspect
    pic = new.shapes.add_picture(snip, Inches(LEFT + (8.90 - pic_w) / 2),
                                 Inches(TOP), width=Inches(pic_w),
                                 height=Inches(pic_h))
    pic.name = name
    tree.remove(pic._element)
    tree.insert(2, pic._element)

    panel, label, body = by['commentary panel'], by['Text 4'], by['Text 5']
    panel.left, panel.top = Inches(LEFT), Inches(strip_top)
    panel.width, panel.height = Inches(8.90), Inches(ph_panel)
    label.left, label.top, label.width = Inches(0.75), Inches(strip_top + 0.08), Inches(3.0)
    extra = by.get('Text 5b')
    if extra is None:
        el = copy.deepcopy(body._element)
        tree.append(el)
        extra = next(sh for sh in new.shapes if sh._element is el)
        extra.name = 'Text 5b'
    for box, x, items in ((body, 0.75, cols[0]), (extra, 5.10, cols[1])):
        box.left, box.top = Inches(x), Inches(strip_top + 0.27)
        box.width, box.height = Inches(4.05), Inches(ph_panel - 0.30)
        write(box, items)
        for para in box.text_frame.paragraphs:
            for r in para.runs:
                r.font.size = Pt(SIZE)

    ns = new.notes_slide
    if ns.notes_text_frame is None:
        ns.shapes._spTree.append(copy.deepcopy(src.notes_slide.notes_placeholder._element))
    ns.notes_text_frame.text = notes

    ids = pres.slides._sldIdLst
    el = ids[-1]
    ids.remove(el)
    ids.insert(at, el)
    return pic_w, pic_h, ph_panel

if __name__ == '__main__':
    BULLETS = [
        'Year to date revenue is $1,624k against a budget of $1,996k, $372k short.',
        'Consulting is most of the gap: $4k billed against a $237k budget.',
        'Integration is $167k short and Video $45k short, both for a third month.',
        'Events is the only department ahead, $420k against a $322k budget.',
        'Support is within $25k of its $743k budget, which is the contract base '
        'behaving as it should.',
    ]
    NOTES = """Revenue against budget for the year to date, by department, taken from the Cost Analysis tab.

    Support $718,587 against a budget of $743,185, short by $24,597 or 3.3%.
    Events $420,121 against $321,957, ahead by $98,163 or 30.5%.
    Video $40,270 against $85,584, short by $45,313 or 52.9%.
    Integration $441,414 against $608,676, short by $167,263 or 27.5%.
    Consulting $4,068 against $236,707, short by $232,640 or 98.3%.

    The company is $371,650 short of budget on revenue, or 18.6%.

    Where the shortfall actually sits. Consulting accounts for $232,640 of it, which is almost two thirds. Integration adds $167,263 and Video $45,313. Events offsets $98,163 of the total and Support is within 3.3% of plan.

    That distribution matters for the reforecast. Support and Events are delivering close to or ahead of what was asked of them. The gap is concentrated in the three departments that were budgeted on a resourcing level which does not exist, and no amount of additional effort from Support or Events closes it.

    Consulting in particular was budgeted $236,707 for the quarter and has billed $4,068. That is not a performance variance, it is a plan built on capacity the business never had.

    Year to date is summed across the closed months rather than read from a stored total, so this chart follows the reporting month on Setup and will extend as the year goes on."""

    p = Presentation(sys.argv[1])
    w, h, strip = add(p, 5, 6, 'FINANCIAL PERFORMANCE',
                      'Revenue against budget, year to date',
                      '/home/user/Claude/deck/snips/rev_ytd.png',
                      BULLETS, NOTES, 'snip rev_ytd')
    p.save(sys.argv[2])
    print(f'slide added as page 7: picture {w:.2f} x {h:.2f}, strip {strip:.2f}')
