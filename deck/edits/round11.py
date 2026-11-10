"""Fix: the two customer slides were showing the old chart.

round10 placed the new pictures but left them named 'snip top10', the name
the old single year chart had. The refresh pass that runs after any snip
rebuild looks each picture up by its name, found top10.png for both, and put
the old chart back on both slides.

Renaming them to match their own snips fixes it and stops it happening again,
because the refresh will now find the right file.
"""
import sys, os
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Emu

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from herdeck import title_of

HERE = os.path.dirname(os.path.abspath(__file__))
SNIP = os.path.join(HERE, '..', 'snips')
# title starts with -> the snip it should carry
WANT = {
    'top ten customers, three year comparison': 'top10_3yr',
    'top ten customers, the figures': 'top10_3yr_tbl',
}
LOGO = os.path.join(HERE, '..', 'assets', 'cts-logo.png')


def relogo(slide):
    """A cloned slide carries the logo's relationship id but not the image
    behind it, so the logo came out as whatever picture held that id. Place it
    from file instead."""
    pic = next((x for x in slide.shapes if x.name == 'cts logo'), None)
    if pic is None or not os.path.exists(LOGO):
        return False
    left, top, width, height = pic.left, pic.top, pic.width, pic.height
    tree = slide.shapes._spTree
    tree.remove(pic._element)
    new = slide.shapes.add_picture(LOGO, left, top, width=width, height=height)
    new.name = 'cts logo'
    return True


def main(src, out):
    pres = Presentation(src)
    fixed = []
    for i, s in enumerate(pres.slides, 1):
        t = title_of(s).lower()
        snip = next((v for k, v in WANT.items() if t.startswith(k)), None)
        if snip is None:
            continue
        pic = next((x for x in s.shapes
                    if 'PICTURE' in str(x.shape_type) and x.name.startswith('snip ')), None)
        panel = next((x for x in s.shapes if x.name == 'commentary panel'), None)
        if pic is None or panel is None:
            print('  slide %d: no picture or no panel' % i)
            continue
        png = os.path.join(SNIP, snip + '.png')
        left = Emu(pic.left).inches
        room_w = 10.0 - 2 * left
        top_lim = 0.94
        bot_lim = Emu(panel.top).inches - 0.12
        im = Image.open(png)
        w, h = room_w, room_w * im.height / im.width
        if h > bot_lim - top_lim:
            h = bot_lim - top_lim
            w = h * im.width / im.height
        tree = s.shapes._spTree
        tree.remove(pic._element)
        new = s.shapes.add_picture(
            png, Inches(left + (room_w - w) / 2),
            Inches(top_lim + (bot_lim - top_lim - h) / 2),
            width=Inches(w), height=Inches(h))
        new.name = 'snip ' + snip
        tree.remove(new._element)
        tree.insert(2, new._element)
        if relogo(s):
            print('  slide %d logo placed from file' % i)
        fixed.append('%d %s -> %s (%.2f x %.2f)' % (i, t[:38], snip, w, h))
    for f in fixed:
        print('  slide %s' % f)
    if len(fixed) != len(WANT):
        print('  WARNING: expected %d slides, fixed %d' % (len(WANT), len(fixed)))
    pres.save(out)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
