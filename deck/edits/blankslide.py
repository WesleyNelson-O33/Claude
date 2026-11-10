"""Add a slide with the heading and the commentary written, and the content
area left empty.

Danica pastes the table in herself. The slide gives her the right heading, the
right section, the commentary already written and the speaker notes, with the
space above the commentary clear to paste into.

The commentary strip is sized for the text it holds, so the empty area above
it is exactly the space the table has to fill.
"""
import copy, sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Inches, Pt
from addslide import balance, write, strip_depth, need, SIZE, TOP, FLOOR, LEFT, LOGO

EYEBROW = 'PEOPLE AND UTILISATION'
TITLE = 'Utilisation Per Employee, First Quarter FY2027'

BULLETS = [
    'Support ran 6,942 chargeable hours against 393 non-chargeable, 95% of hours worked.',
    'Production billed 954 hours against 1,150 non-chargeable, so over half its worked time was not billed.',
    'Consulting is the same shape on a smaller base, 179 chargeable against 280.',
    'Leave took 1,258 hours across the quarter, 543 of it in Production.',
    'Across CTS the quarter ran 8,074 chargeable hours against 1,823 non-chargeable.',
]

NOTES = (
    'Utilisation by person for the quarter.\n\n'
    'Support is carrying the quarter on utilisation the same way it carries it on '
    'profit. 6,942 chargeable hours against 393 non-chargeable is 95 per cent of the '
    'hours its people actually worked.\n\n'
    'Production is the opposite. 954 chargeable against 1,150 non-chargeable, so more '
    'than half of what its people worked was not billed to anything.\n\n'
    'Consulting is the same shape on a much smaller base, 179 against 280.\n\n'
    'Leave took 1,258 hours across the quarter. 543 of that sits in Production, on a '
    'far smaller headcount than Support, which is part of why its chargeable share '
    'reads low.\n\n'
    'Across CTS the quarter ran 8,074 chargeable hours against 1,823 non-chargeable.'
)


def main(src, out, at):
    pres = Presentation(src)
    before = len(pres.slides)
    template = pres.slides[5]
    new = pres.slides.add_slide(template.slide_layout)
    for sh in list(new.shapes):
        sh._element.getparent().remove(sh._element)
    for sh in template.shapes:
        new.shapes._spTree.append(copy.deepcopy(sh._element))

    by = {sh.name: sh for sh in new.shapes}
    tree = new.shapes._spTree
    by['Text 0'].text_frame.paragraphs[0].runs[0].text = EYEBROW
    by['Text 1'].text_frame.paragraphs[0].runs[0].text = TITLE

    # every picture comes off, then the logo goes back from file
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

    cols = list(balance(BULLETS, 4.05))
    ph_panel = strip_depth(max(need(c, 4.05) for c in cols if c) + 0.30)
    strip_top = FLOOR - ph_panel

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
        ns.shapes._spTree.append(copy.deepcopy(template.notes_slide.notes_placeholder._element))
    ns.notes_text_frame.text = NOTES

    ids = pres.slides._sldIdLst
    el = ids[-1]
    ids.remove(el)
    ids.insert(int(at), el)
    pres.save(out)
    print('  added at index %s, paste area %.2f wide by %.2f high starting %.2f down'
          % (at, 8.90, strip_top - 0.12 - TOP, TOP))
    print('%d slides in, %d out' % (before, len(pres.slides)))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3])
