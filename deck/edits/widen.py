"""Re-place the pictures on the pages whose chart was drawn nearly square.

Those charts have been widened in the pack, so the snip that comes out of it is
a wider picture and fills the page the way the rest of the deck does. Nothing
else on the page changes: the existing commentary is read off the slide and
written back word for word, and the speaker notes are left as they are.
"""
import sys
from pptx import Presentation
from pptx.util import Emu
from addslide import relay

SNIP = '/home/user/Claude/deck/snips/'

# page number -> snip name, as the picture on that page is called
PAGES = {
    17: 'q4_fullyear',
    18: 'q4_quarter',
    19: 'q4_bridge',
    33: 'cost_dist',
    34: 'other_opex',
    35: 'labour_mix',
    38: 'top10',
    39: 'dso_calc',
    40: 'dso_who',
    45: 'q4_perhead',
    46: 'q4_ratios',
}


def read_bullets(slide):
    """The commentary as it stands, left column then right, in order."""
    out = []
    for key in ('Text 5', 'Text 5b', 'Text 6'):
        box = next((sh for sh in slide.shapes if sh.name == key), None)
        if box is None:
            continue
        for para in box.text_frame.paragraphs:
            text = ''.join(r.text for r in para.runs).strip()
            if text:
                out.append(text)
    return out


def main(src, out):
    pres = Presentation(src)
    for page, snip in PAGES.items():
        slide = pres.slides[page - 1]
        name = 'snip ' + snip
        pic = next((sh for sh in slide.shapes if sh.name == name), None)
        if pic is None:
            print(f'  page {page}: no picture called {name!r}, skipped')
            continue
        was = Emu(pic.width).inches
        bullets = read_bullets(slide)
        notes = slide.notes_slide.notes_text_frame.text
        w, h, strip = relay(slide, SNIP + snip + '.png', bullets, notes, name)
        print(f'  page {page:>3} {snip:<12} {was:>5.2f} -> {w:>5.2f} wide, '
              f'{h:.2f} tall, strip {strip:.2f}')
    pres.save(out)
    print(f'{len(pres.slides)} slides saved')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
