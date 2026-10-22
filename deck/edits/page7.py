"""Page 7: month by month instead of a year to date ranking."""
import copy, sys
from PIL import Image
from pptx import Presentation
from pptx.text.text import _Paragraph
from pptx.util import Inches, Pt

SNIP = '/home/user/Claude/deck/snips/rev_ytd.png'
TITLE = 'Revenue against budget, month by month'
TOP, FLOOR, LEFT = 1.22, 5.06, 0.55
SIZE, LEAD, GAP = 10.0, 12.5, 4.0

BULLETS = [
    'July billed $459k against a $609k budget, $150k short.',
    'August billed $603k against a $973k budget, $370k short.',
    'September billed $562k against a $414k budget, $148k ahead.',
    'Year to date $1,624k against $1,996k, a shortfall of $372k or 18.6%.',
    'Two thirds of the shortfall is August alone, where the budget asked for '
    'more than the business has ever billed in a month.',
]
NOTES = """Revenue against budget for each closed month, taken from the Cost Analysis tab.

July: $459,049 against a budget of $609,277, short by $150,228 or 24.7%.
August: $603,467 against $973,168, short by $369,701 or 38.0%.
September: $561,945 against $413,665, ahead by $148,280 or 35.8%.

Year to date: $1,624,460 against $1,996,110, short by $371,650 or 18.6%.

The monthly shape matters more than the quarter total. Two thirds of the shortfall is August alone, where the budget asked for $973,168. That is more than the business has billed in any month on record, so the variance reflects the plan as much as the trading.

September is the only month ahead of budget, and that came from a single Integration installation completing at a 10.1% gross margin, so it added revenue without adding much profit.

Read alongside the previous page, which shows which department billed each month, the two slides together explain the quarter: a steady Support base, one large project landing in September, and a budget that front-loaded the year.

The chart takes the closed months only and lengthens as the year goes on. The year to date pair sits on the table beneath it rather than on the chart, because a column five times the height of the others would flatten the months."""

def lines(b, w):
    cpl = max(12, int(32 * (10 / SIZE) * (w / 2.20)))
    return sum(max(1, -(-len(t) // cpl)) for t in b)

def need(b, w):
    return (lines(b, w) * LEAD + (len(b) - 1) * GAP) / 72 / 0.94

def balance(b, w):
    best, where = None, 1
    for k in range(1, len(b)):
        worst = max(lines(b[:k], w), lines(b[k:], w))
        if best is None or worst < best:
            best, where = worst, k
    return b[:where], b[where:]

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
s = p.slides[6]
by = {sh.name: sh for sh in s.shapes}
tree = s.shapes._spTree
by['Text 1'].text_frame.paragraphs[0].runs[0].text = TITLE

old = next(sh for sh in s.shapes
           if 'PICTURE' in str(sh.shape_type) and sh.name != 'cts logo')
tree.remove(old._element)

cols = list(balance(BULLETS, 4.05))
ph_panel = max(need(c, 4.05) for c in cols if c) + 0.30
strip_top = FLOOR - ph_panel

im = Image.open(SNIP)
aspect = im.width / im.height
pic_h = min(strip_top - 0.12 - TOP, 8.90 / aspect)
pic_w = pic_h * aspect
pic = s.shapes.add_picture(SNIP, Inches(LEFT + (8.90 - pic_w) / 2), Inches(TOP),
                           width=Inches(pic_w), height=Inches(pic_h))
pic.name = 'snip rev_month_budget'
tree.remove(pic._element)
tree.insert(2, pic._element)

panel, label = by['commentary panel'], by['Text 4']
panel.left, panel.top = Inches(LEFT), Inches(strip_top)
panel.width, panel.height = Inches(8.90), Inches(ph_panel)
label.left, label.top, label.width = Inches(0.75), Inches(strip_top + 0.08), Inches(3.0)
for box, x, items in ((by['Text 5'], 0.75, cols[0]), (by['Text 5b'], 5.10, cols[1])):
    box.left, box.top = Inches(x), Inches(strip_top + 0.27)
    box.width, box.height = Inches(4.05), Inches(ph_panel - 0.30)
    write(box, items)
    for para in box.text_frame.paragraphs:
        for r in para.runs:
            r.font.size = Pt(SIZE)

s.notes_slide.notes_text_frame.text = NOTES
p.save(sys.argv[2])
print(f'page 7 rebuilt: picture {pic_w:.2f} x {pic_h:.2f}, strip {ph_panel:.2f}')
