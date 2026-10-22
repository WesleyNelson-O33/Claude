"""One frame for every page that carries its commentary at the bottom.

The pictures had been dragged to size by hand, which made seven of them wider or
taller than the picture really is. Page 13 was the worst, 26 per cent out. On a
chart that widens the bars and stretches the words inside them.

So every one of these pages now uses the same box: it starts at the left margin,
runs the full 8.90 inches of the content width, starts 1.10 inches down and stops
a tenth of an inch above the commentary panel. Each picture is scaled to fit that
box at its own true shape and centred in it, so nothing is stretched.

The charts that sit alone on one of these pages were redrawn in the pack at the
shape of the box, so they fill the width on their own rather than being pulled
out to it.

Page 7 carries two charts, so the box is split into two halves with a gap.

Pages whose commentary sits down the right hand side are a different layout and
are left alone.
"""
import sys
from PIL import Image
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

SNIP = '/home/user/Claude/deck/snips'
LEFT, WIDTH, TOP, CLEAR, GAP = 0.55, 8.90, 1.10, 0.10, 0.20
TITLE_TOP, TITLE_H = 0.70, 0.40      # the box is middle anchored, so the words
                                     # sit at 0.90 and stay clear of the picture

# Only the charts that were redrawn in the pack are put back from file. The
# others keep the picture already on the slide: pages 5, 11 and 14 to 18 carry a
# cropped version with rows dropped out of it, which is not the plain snip, and
# replacing those from file would undo the crop.
REDRAWN = {
    'snip rev_month': 'rev_month', 'snip rev_q1': 'xero_qtr',
    'snip xero_hist': 'xero_hist', 'snip xero_cum': 'xero_cum',
    'snip rev_by_dept': 'rev_by_dept',
}


def place(slide, sh, x, y, w, h):
    """Fit the picture into the box at its own shape and centre it there.

    A redrawn chart is put back from file so the deck and the pack stay in step.
    Everything else keeps the picture it already has and is only resized, because
    the image on the slide is not always the plain snip."""
    name = sh.name
    path = f'{SNIP}/{REDRAWN[name]}.png' if name in REDRAWN else None
    if path:
        im = Image.open(path)
        aspect = im.width / im.height
    else:
        iw, ih = sh.image.size
        aspect = iw / ih
    ph = min(h, w / aspect)
    pw = ph * aspect
    left, top = Inches(x + (w - pw) / 2), Inches(y + (h - ph) / 2)
    if path:
        tree = slide.shapes._spTree
        tree.remove(sh._element)
        pic = slide.shapes.add_picture(path, left, top,
                                       width=Inches(pw), height=Inches(ph))
        pic.name = name
        tree.remove(pic._element)
        tree.insert(2, pic._element)
    else:
        sh.left, sh.top = left, top
        sh.width, sh.height = Inches(pw), Inches(ph)
    return pw, ph


p = Presentation(sys.argv[1])
for i, slide in enumerate(p.slides):
    panel = next((sh for sh in slide.shapes if sh.name == 'commentary panel'), None)
    pics = [sh for sh in slide.shapes
            if 'PICTURE' in str(sh.shape_type) and sh.name != 'cts logo']
    if panel is None or not pics:
        continue
    # the title: centred, and in a box the same depth on every page, so the
    # words sit at the same height all the way through. Every title is one line
    # at this size in this width, so a shallow box cannot clip one.
    title = next((sh for sh in slide.shapes if sh.name == 'Text 1'), None)
    if title is not None:
        title.top, title.height = Inches(TITLE_TOP), Inches(TITLE_H)
        for para in title.text_frame.paragraphs:
            para.alignment = PP_ALIGN.CENTER

    panel_top = panel.top / 914400
    if panel_top < 3.0:                       # commentary down the right, not ours
        print(f'p{i+1:>2} left alone, commentary on the right')
        continue

    box_h = panel_top - CLEAR - TOP
    if len(pics) == 1:
        name = pics[0].name
        w, h = place(slide, pics[0], LEFT, TOP, WIDTH, box_h)
        print(f'p{i+1:>2} {name:24} {w:.2f}x{h:.2f} in a {WIDTH:.2f}x{box_h:.2f} box')
    else:
        half = (WIDTH - GAP) / 2
        for k, sh in enumerate(sorted(pics, key=lambda s: s.left)):
            name = sh.name
            x = LEFT + k * (half + GAP)
            w, h = place(slide, sh, x, TOP, half, box_h)
            print(f'p{i+1:>2} {name:24} {w:.2f}x{h:.2f} in a {half:.2f}x{box_h:.2f} half')

p.save(sys.argv[2])
