"""Refresh every chart image in the deck from the snips, in place.

Her positions are kept exactly: the top left and the width stay, only the
height follows the new aspect, so nothing she has arranged moves.
"""
import sys, os
from PIL import Image
from pptx import Presentation

SNIP = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'snips')


def main(src, out):
    pres = Presentation(src)
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
    pres.save(out)
    print('  %d chart images refreshed, %d slides' % (n, len(pres.slides)))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
