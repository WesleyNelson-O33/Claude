"""Put the current snip PNGs back onto the slides.

Changing a chart in the pack changes the PNG on disk, but the deck holds its
own copy of every picture, so the slides keep showing the old one until they
are re-placed. This walks the deck and swaps each one.

Geometry is recomputed, because a chart that loses its title gets shorter and
its aspect ratio changes. The picture is sized to the gap between the title and
the commentary strip, exactly as addslide.add() sizes it, and centred.

A slide carrying more than one chart is left alone. Those have a hand-built
side by side layout that this would flatten.
"""
import copy, sys, os
from PIL import Image
from pptx import Presentation
from pptx.util import Inches

SNIPDIR = '/home/user/Claude/deck/snips/'
TOP, FLOOR, LEFT, SPAN = 1.22, 5.06, 0.55, 8.90


def emu_in(v):
    return v / 914400.0


def main(src, out):
    pres = Presentation(src)
    swapped = skipped = missing = 0
    for i, s in enumerate(pres.slides, 1):
        pics = [sh for sh in s.shapes
                if 'PICTURE' in str(sh.shape_type) and (sh.name or '').startswith('snip ')]
        if not pics:
            continue
        if len(pics) > 1:
            print('  %-3d left alone, %d charts on the slide' % (i, len(pics)))
            skipped += 1
            continue

        pic = pics[0]
        name = pic.name
        png = os.path.join(SNIPDIR, name[5:] + '.png')
        if not os.path.exists(png):
            print('  %-3d %-18s PNG MISSING, left as it is' % (i, name[5:]))
            missing += 1
            continue

        panel = next((sh for sh in s.shapes if sh.name == 'commentary panel'), None)
        tree = s.shapes._spTree
        old_box = (pic.left, pic.top, pic.width, pic.height)
        tree.remove(pic._element)

        im = Image.open(png)
        aspect = im.width / im.height
        if panel is not None:
            strip_top = emu_in(panel.top)
            pic_h = min(strip_top - 0.12 - TOP, SPAN / aspect)
            pic_w = pic_h * aspect
            left, top = LEFT + (SPAN - pic_w) / 2, TOP
        else:
            # no strip on this one, so keep the box it had and fit inside it
            left, top = emu_in(old_box[0]), emu_in(old_box[1])
            box_w, box_h = emu_in(old_box[2]), emu_in(old_box[3])
            pic_h = min(box_h, box_w / aspect)
            pic_w = pic_h * aspect

        new = s.shapes.add_picture(png, Inches(left), Inches(top),
                                   width=Inches(pic_w), height=Inches(pic_h))
        new.name = name
        tree.remove(new._element)
        tree.insert(2, new._element)
        print('  %-3d %-18s %.2f x %.2f' % (i, name[5:], pic_w, pic_h))
        swapped += 1

    pres.save(out)
    print('\n%d swapped, %d left alone, %d missing a PNG' % (swapped, skipped, missing))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
