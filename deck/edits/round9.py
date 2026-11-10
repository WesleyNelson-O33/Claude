"""Round nine: the utilisation schedule comes from the pack now.

The table was built in a workbook of its own and snipped from there, which
made it the one thing in the deck not read from the controller pack. It is in
the pack's Utilisation sheet now, carrying Danica's own conditional
formatting, the green data bar across Over / under.

The new table is wider and shorter than the old one, so the picture is
refitted: full width of the space left of the commentary panel, height by its
own aspect, and centred in the room it has so there is no band of white
between it and the panel.
"""
import sys, os
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Emu

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from herdeck import title_of

HERE = os.path.dirname(os.path.abspath(__file__))
SNIP = os.path.join(HERE, '..', 'snips')
GAP = 0.20                                  # breathing room beside the panel


def refit(slide, name):
    png = os.path.join(SNIP, name[5:] + '.png')
    if not os.path.exists(png):
        return None
    pic = next((s for s in slide.shapes if s.name == name), None)
    panel = next((s for s in slide.shapes if s.name == 'commentary panel'), None)
    if pic is None or panel is None:
        return None
    left = Emu(pic.left).inches
    room_w = Emu(panel.left).inches - left - GAP
    top_lim, bot_lim = Emu(panel.top).inches, Emu(panel.top).inches + Emu(panel.height).inches
    im = Image.open(png)
    w = room_w
    h = w * im.height / im.width
    if h > bot_lim - top_lim:
        h = bot_lim - top_lim
        w = h * im.width / im.height
    tree = slide.shapes._spTree
    tree.remove(pic._element)
    new = slide.shapes.add_picture(
        png, Inches(left + (room_w - w) / 2), Inches(top_lim + (bot_lim - top_lim - h) / 2),
        width=Inches(w), height=Inches(h))
    new.name = name
    tree.remove(new._element)
    tree.insert(2, new._element)
    return w, h


def main(src, out):
    pres = Presentation(src)
    for i, s in enumerate(pres.slides, 1):
        if 'utilisation per employee' not in title_of(s).lower():
            continue
        got = refit(s, 'snip util_q1')
        if got is None:
            print('  slide %d: could not refit the table' % i)
            return
        print('  slide %d table refitted to %.2f x %.2f inches' % (i, got[0], got[1]))
        break
    pres.save(out)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
