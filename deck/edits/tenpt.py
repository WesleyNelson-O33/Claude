"""Every commentary bullet at 10 point, with the panels resized to hold it.

Ten point is larger than the panels were built for, so each one is re-measured:
the bottom strips grow downward into the space the picture gives up, and the
side panels widen. The picture is then refitted to whatever is left.
"""
import copy, io, sys
from PIL import Image
from pptx import Presentation
from pptx.text.text import _Paragraph
from pptx.util import Inches, Pt

SIZE = 10.0
LEAD = 12.5
GAP = 4.0
FLOOR = 5.06
RIGHT = 9.45
LEFT = 0.55

def lines(bullets, w_in):
    """Wrapped lines for a column of a given width, Calibri at 10 point."""
    cpl = max(12, int(32 * (10 / SIZE) * (w_in / 2.20)))
    return sum(max(1, -(-len(t) // cpl)) for t in bullets)

def need(bullets, w_in):
    return (lines(bullets, w_in) * LEAD + (len(bullets) - 1) * GAP) / 72 / 0.94

def bullets_of(box):
    return [p.text for p in box.text_frame.paragraphs if p.text.strip()]

def balance(bullets, w_in):
    """Split the bullets between two columns so neither runs deeper than it
    has to. A fixed split leaves one column three lines longer than the
    other, and the strip has to be as tall as the worst of the two."""
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


def set_size(box):
    for para in box.text_frame.paragraphs:
        for r in para.runs:
            r.font.size = Pt(SIZE)

def fix(slide):
    by = {sh.name: sh for sh in slide.shapes}
    panel = by.get('commentary panel')
    if panel is None:
        return None
    pic = next((sh for sh in slide.shapes
                if 'PICTURE' in str(sh.shape_type) and sh.name != 'cts logo'), None)
    if pic is None:
        return None
    im = Image.open(io.BytesIO(pic.image.blob))
    aspect = im.width / im.height
    boxes = [by[n] for n in ('Text 5', 'Text 5b', 'Text 6') if n in by]
    cols = [bullets_of(b) for b in boxes]
    top = pic.top / 914400
    wide = panel.width / 914400 > 5

    if wide:                       # commentary in a strip under the picture
        flat = [b for c in cols for b in c]
        cols = list(balance(flat, 4.05))
        h_text = max(need(c, 4.05) for c in cols if c)
        ph = h_text + 0.30                    # panel height
        strip_top = FLOOR - ph
        pic_h = min(strip_top - 0.12 - top, 8.90 / aspect)
        pic_w = pic_h * aspect
        pic.left, pic.width, pic.height = Inches(LEFT + (8.90 - pic_w) / 2), \
            Inches(pic_w), Inches(pic_h)
        panel.left, panel.top = Inches(LEFT), Inches(strip_top)
        panel.width, panel.height = Inches(8.90), Inches(ph)
        by['Text 4'].left, by['Text 4'].top = Inches(0.75), Inches(strip_top + 0.08)
        for box, x, items in zip(boxes, (0.75, 5.10), cols):
            box.left, box.top = Inches(x), Inches(strip_top + 0.27)
            box.width, box.height = Inches(4.05), Inches(ph - 0.30)
            write(box, items)
            set_size(box)
        return ('strip', pic_w, pic_h, ph)

    # commentary in a panel beside the picture: widen it until the text fits
    flat = [b for c in cols for b in c]
    avail_h = FLOOR - (panel.top / 914400) - 0.50
    w = 2.60
    while w < 3.60 and need(flat, w - 0.40) > avail_h:
        w += 0.10
    panel.left, panel.width = Inches(RIGHT - w), Inches(w)
    by['Text 4'].left, by['Text 4'].width = Inches(RIGHT - w + 0.20), Inches(w - 0.40)
    for box in boxes:
        box.left, box.width = Inches(RIGHT - w + 0.20), Inches(w - 0.40)
        box.height = Inches(avail_h)
        set_size(box)
    box_w = (RIGHT - w - 0.25) - LEFT
    pic_h = min(box_w / aspect, FLOOR - top - 0.05)
    pic_w = pic_h * aspect
    pic.left, pic.width, pic.height = Inches(LEFT), Inches(pic_w), Inches(pic_h)
    return ('panel', pic_w, pic_h, w)

p = Presentation(sys.argv[1])
for i, s in enumerate(p.slides, 1):
    out = fix(s)
    if out:
        kind, w, h, extra = out
        print(f'slide {i:>3} {kind:6s} picture {w:.2f} x {h:.2f}   '
              f'{"strip" if kind == "strip" else "panel"} {extra:.2f}')
p.save(sys.argv[2])
