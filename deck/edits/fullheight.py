"""Give a dense table slide the whole slide instead of the standard band.

The usual chart page leaves about two and a quarter inches between the heading
and the commentary strip. That band is roughly four times as wide as it is
tall, so anything close to square has to shrink to about a third of the slide
width to fit, and becomes unreadable.

A full profit and loss is close to square. So on the slides named here the
commentary strip comes off and the table takes the full height. What the
commentary said moves into the speaker notes, where it is read out anyway.
"""
import sys, os
from pptx import Presentation
from pptx.util import Inches
from PIL import Image

SNIPDIR = '/home/user/Claude/deck/snips/'
TOP, FLOOR, LEFT, SPAN = 1.22, 5.06, 0.55, 8.90
TITLES = {'profit and loss summary, first quarter fy2027'}
STRIP = {'commentary panel', 'Text 4', 'Text 5', 'Text 5b'}


def title_of(slide):
    for sh in slide.shapes:
        if sh.name == 'Text 1' and sh.has_text_frame:
            return sh.text_frame.text.strip().split('\n')[0]
    return ''


def main(src, out):
    pres = Presentation(src)
    done = 0
    for i, s in enumerate(pres.slides, 1):
        if title_of(s).lower() not in TITLES:
            continue
        tree = s.shapes._spTree

        # fold the commentary into the notes before dropping it
        lines = []
        for name in ('Text 5', 'Text 5b'):
            sh = next((x for x in s.shapes if x.name == name), None)
            if sh is not None and sh.has_text_frame:
                lines += [p.text.strip() for p in sh.text_frame.paragraphs if p.text.strip()]
        if lines:
            ns = s.notes_slide
            ns.notes_text_frame.text = (ns.notes_text_frame.text.rstrip()
                                        + '\n\n' + '\n'.join(lines)).strip()

        for sh in list(s.shapes):
            if sh.name in STRIP:
                tree.remove(sh._element)

        pic = next((sh for sh in s.shapes
                    if 'PICTURE' in str(sh.shape_type) and sh.name.startswith('snip ')), None)
        if pic is None:
            continue
        name = pic.name
        png = os.path.join(SNIPDIR, name[5:] + '.png')
        tree.remove(pic._element)
        im = Image.open(png)
        aspect = im.width / im.height
        room = FLOOR - TOP
        pic_h = min(room, SPAN / aspect)
        pic_w = pic_h * aspect
        new = s.shapes.add_picture(png, Inches(LEFT + (SPAN - pic_w) / 2),
                                   Inches(TOP + (room - pic_h) / 2),
                                   width=Inches(pic_w), height=Inches(pic_h))
        new.name = name
        tree.remove(new._element)
        tree.insert(2, new._element)
        print('  %-3d %-46s %.2f x %.2f' % (i, title_of(s)[:46], pic_w, pic_h))
        done += 1
    pres.save(out)
    print('%d slides given the full height' % done)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
