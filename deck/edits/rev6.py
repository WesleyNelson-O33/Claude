"""Page 6: the budget line added, and the commentary moved to the bottom so it
reads the same as the rest of the deck."""
import copy, sys
from PIL import Image
from pptx import Presentation
from pptx.text.text import _Paragraph
from pptx.util import Inches, Pt

SNIP = '/home/user/Claude/deck/snips/rev_month.png'
TOP, FLOOR, LEFT = 1.22, 5.06, 0.55
SIZE, LEAD, GAP = 10.0, 12.5, 4.0

BULLETS = [
    'Revenue was $459k in July, $603k in August and $562k in September, '
    '$1,624k for the quarter.',
    'Budget for the same three months was $1,996k, so the quarter is $372k short.',
    'July and August were both well under budget. September came in $148k above it.',
    'Support is the steady base at $235k, $240k and $243k.',
    "September's movement is Integration at $235k against $94k in August, on "
    'project completion.',
]
NOTES = """Revenue by month and department for the closed quarter, with the budget for the same months, taken from the Cost Analysis tab.

Actual revenue was $459,049 in July, $603,467 in August and $561,945 in September, $1,624,460 for the quarter.

Budget for those months was $609,277, $973,168 and $413,665, being $1,996,110 for the quarter. The quarter is therefore $371,650 short of budget.

The monthly picture is not what the quarter total suggests. July came in $150,228 under budget and August $369,701 under, while September came in $148,280 above. Two thirds of the quarter's shortfall sits in August alone, where the budget asked for $973,168, more than the business has ever billed in a month.

By department for September: Support $242,948, Events $73,304, Video $7,303, Integration $234,858 and Consulting $3,531.

Three things the mix shows that a total does not.

Support is the base of the business and it barely moves: $235,403, $240,236 and $242,948 across the three months. That is recurring contract income behaving as recurring income should.

Integration is the volatility. It billed $112,803 in July, $93,752 in August and $234,858 in September, and the September figure is a single installation completing. It is better read as timing than as growth, particularly as the margin retained on that work is 10.1%.

Events moved from $241,371 in August to $73,304 in September, the largest single movement in the quarter, and the reason the company total fell while Integration rose.

The chart lengthens by one column each month as the pack closes a month, so it will show the full year by June without being rebuilt."""

def lines(bullets, w_in):
    cpl = max(12, int(32 * (10 / SIZE) * (w_in / 2.20)))
    return sum(max(1, -(-len(t) // cpl)) for t in bullets)

def need(bullets, w_in):
    return (lines(bullets, w_in) * LEAD + (len(bullets) - 1) * GAP) / 72 / 0.94

def balance(bullets, w_in):
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

p = Presentation(sys.argv[1])
s = p.slides[5]
by = {sh.name: sh for sh in s.shapes}
tree = s.shapes._spTree

old = next(sh for sh in s.shapes
           if 'PICTURE' in str(sh.shape_type) and sh.name != 'cts logo')
tree.remove(old._element)

# The logo was deep-copied from page 16 when this slide was created, which
# carried its relationship id across but not the image behind it, so it was
# pointing at whatever part held that id on the new slide. Put it back properly.
LOGO = '/home/user/Claude/deck/assets/cts-logo.png'
stale = next((sh for sh in s.shapes if sh.name == 'cts logo'), None)
if stale is not None:
    geom = (stale.left, stale.top, stale.width, stale.height)
    tree.remove(stale._element)
    mark = s.shapes.add_picture(LOGO, geom[0], geom[1], width=geom[2], height=geom[3])
    mark.name = 'cts logo'

cols = list(balance(BULLETS, 4.05))
ph_panel = max(c and need(c, 4.05) or 0 for c in cols) + 0.30
strip_top = FLOOR - ph_panel

im = Image.open(SNIP)
aspect = im.width / im.height
pic_h = min(strip_top - 0.12 - TOP, 8.90 / aspect)
pic_w = pic_h * aspect
pic = s.shapes.add_picture(SNIP, Inches(LEFT + (8.90 - pic_w) / 2), Inches(TOP),
                           width=Inches(pic_w), height=Inches(pic_h))
pic.name = 'snip rev_month'
tree.remove(pic._element)
tree.insert(2, pic._element)

panel, label, body = by['commentary panel'], by['Text 4'], by['Text 5']
panel.left, panel.top = Inches(LEFT), Inches(strip_top)
panel.width, panel.height = Inches(8.90), Inches(ph_panel)
label.left, label.top = Inches(0.75), Inches(strip_top + 0.08)
label.width = Inches(3.0)

second = copy.deepcopy(body._element)
tree.append(second)
body2 = next(sh for sh in s.shapes if sh._element is second)
body2.name = 'Text 5b'
for box, x, items in ((body, 0.75, cols[0]), (body2, 5.10, cols[1])):
    box.left, box.top = Inches(x), Inches(strip_top + 0.27)
    box.width, box.height = Inches(4.05), Inches(ph_panel - 0.30)
    write(box, items)
    for para in box.text_frame.paragraphs:
        for r in para.runs:
            r.font.size = Pt(SIZE)

s.notes_slide.notes_text_frame.text = NOTES
p.save(sys.argv[2])
print(f'page 6 rebuilt: picture {pic_w:.2f} x {pic_h:.2f}, strip {ph_panel:.2f}')
