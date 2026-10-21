"""A revenue slide, built from page 16 so it inherits every setting, and
inserted straight after the P&L summary income page."""
import copy, sys
from PIL import Image
from pptx import Presentation
from pptx.text.text import _Paragraph
from pptx.util import Inches, Pt

SNIP = '/home/user/Claude/deck/snips/rev_month.png'
TITLE = 'Revenue by month and department'
SOURCE = ('Snipped from the Financial Controller Pack, Cost Analysis tab, '
          'revenue by month and department, the closed months.')
BULLETS = [
    'Revenue was $459k in July, $603k in August and $562k in September.',
    'Support is the steady base: $235k, $240k and $243k across the three months.',
    "September's movement is Integration at $235k against $94k in August, on "
    'project completion.',
    'Events fell from $241k in August to $73k in September, the largest movement '
    'in the quarter.',
    'Video and Consulting contributed $11k in September against a $41k combined '
    'budget.',
]
NOTES = """Revenue by month and department for the closed quarter, taken from the Cost Analysis tab.

July $459,049, August $603,467 and September $561,945, giving $1,624,460 for the quarter against a budget of $1,996,110.

By department for September: Support $242,948, Events $73,304, Video $7,303, Integration $234,858 and Consulting $3,531.

Three things the mix shows that a total does not.

Support is the base of the business and it barely moves: $235,403, $240,236 and $242,948 across the three months. That is recurring contract income behaving as recurring income should.

Integration is the volatility. It billed $112,803 in July, $93,752 in August and $234,858 in September, and the September figure is a single installation completing. It is better read as timing than as growth, particularly as the margin retained on that work is 10.1%.

Events moved from $241,371 in August to $73,304 in September, the largest single movement in the quarter, and the reason the company total fell while Integration rose.

Video and Consulting contributed $10,834 between them in September against a combined budget of $40,548, which is the third consecutive month both have been short.

The chart lengthens by one column each month as the pack closes a month, so it will show the full year by June without being rebuilt."""

def clone(src, dst_pres):
    """A new slide on the same layout, carrying copies of every shape."""
    new = dst_pres.slides.add_slide(src.slide_layout)
    for sh in list(new.shapes):
        sh._element.getparent().remove(sh._element)
    for sh in src.shapes:
        new.shapes._spTree.append(copy.deepcopy(sh._element))
    return new

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

p = Presentation(sys.argv[1])
new = clone(p.slides[15], p)                    # page 16, a chart on a side panel
by = {sh.name: sh for sh in new.shapes}

by['Text 1'].text_frame.paragraphs[0].runs[0].text = TITLE
by['Text 0'].text_frame.paragraphs[0].runs[0].text = 'FINANCIAL PERFORMANCE'

# page 16 carries the widest panel in the deck; this slide's commentary is
# shorter, so the panel comes in and the chart takes the width back
PANEL_W, RIGHT = 3.10, 9.45
for name in ('commentary panel',):
    by[name].left, by[name].width = Inches(RIGHT - PANEL_W), Inches(PANEL_W)
for name in ('Text 4', 'Text 5'):
    by[name].left = Inches(RIGHT - PANEL_W + 0.20)
    by[name].width = Inches(PANEL_W - 0.40)

old = next(sh for sh in new.shapes
           if 'PICTURE' in str(sh.shape_type) and sh.name != 'cts logo')
top = old.top
tree = new.shapes._spTree
tree.remove(old._element)
im = Image.open(SNIP)
aspect = im.width / im.height
pw = (RIGHT - PANEL_W - 0.25) - 0.55
ph = pw / aspect
pic = new.shapes.add_picture(SNIP, Inches(0.55), top, width=Inches(pw), height=Inches(ph))
pic.name = 'snip rev_month'
tree.remove(pic._element)
tree.insert(2, pic._element)

write(by['Text 5'], BULLETS)
for para in by['Text 5'].text_frame.paragraphs:
    for r in para.runs:
        r.font.size = Pt(10)

# the fresh notes slide has no body placeholder, so borrow one from page 16
src_ns, dst_ns = p.slides[15].notes_slide, new.notes_slide
if dst_ns.notes_text_frame is None:
    body = src_ns.notes_placeholder._element
    dst_ns.shapes._spTree.append(copy.deepcopy(body))
dst_ns.notes_text_frame.text = NOTES

# move it in behind the P&L summary income page
ids = p.slides._sldIdLst
el = ids[-1]
ids.remove(el)
ids.insert(5, el)

p.save(sys.argv[2])
print('revenue slide inserted as page 6')
