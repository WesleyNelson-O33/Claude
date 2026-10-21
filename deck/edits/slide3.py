"""Slide 3: fit a sixth card for the company total, and reword the strip."""
import copy, sys
from pptx import Presentation
from pptx.util import Emu, Inches

SRC = sys.argv[1]
OUT = sys.argv[2]
NAVY, WHITE, SOFT = 0x242C39, 0xFFFFFF, 0xE3E6EA

p = Presentation(SRC)
s = p.slides[2]
by = {sh.name: sh for sh in s.shapes}

CARDS = ['Support', 'Events', 'Video', 'Integration', 'Consulting']
# five cards of 1.62 with 0.17 gaps do not leave room for a sixth, so the row
# is re-pitched: six cards of 1.3833 with 0.12 gaps still spans 0.55 to 9.45
W, GAP, X0, PAD = 1.38333, 0.12, 0.55, 0.12
left = lambda i: X0 + i * (W + GAP)

# the six text boxes that sit on each card, in order down the card
ROWS = {'Support': [3, 4, 5, 6, 7, 8], 'Events': [10, 11, 12, 13, 14, 15],
        'Video': [17, 18, 19, 20, 21, 22], 'Integration': [24, 25, 26, 27, 28, 29],
        'Consulting': [31, 32, 33, 34, 35, 36]}

for i, name in enumerate(CARDS):
    card = by[f'{name} card']
    card.left, card.width = Inches(left(i)), Inches(W)
    for n in ROWS[name]:
        t = by[f'Text {n}']
        t.left, t.width = Inches(left(i) + PAD), Inches(W - 2 * PAD)

# the sixth card, copied off Consulting so it inherits every setting
tree = s.shapes._spTree
src_card = by['Consulting card']
new_card = copy.deepcopy(src_card._element)
tree.append(new_card)
total_card = [sh for sh in s.shapes if sh._element is new_card][0]
total_card.name = 'Total card'
total_card.left, total_card.width = Inches(left(5)), Inches(W)
total_card.fill.solid()
total_card.fill.fore_color.rgb = __import__('pptx.dml.color', fromlist=['RGBColor']).RGBColor(0x24, 0x2C, 0x39)
total_card.line.fill.background()

TEXT = ['TOTAL', '$562k', 'revenue', '33% gross margin',
        '$(41)k after overhead', 'Ahead of budget by $148k']
for j, n in enumerate(ROWS['Consulting']):
    src = by[f'Text {n}']
    el = copy.deepcopy(src._element)
    tree.append(el)
    box = [sh for sh in s.shapes if sh._element is el][0]
    box.name = f'Total text {j}'
    box.left, box.width = Inches(left(5) + PAD), Inches(W - 2 * PAD)
    para = box.text_frame.paragraphs[0]
    for r in para.runs[1:]:
        r._r.getparent().remove(r._r)
    para.runs[0].text = TEXT[j]

STRIP = ('September revenue exceeded budget by $148k, while EBITDA remained a '
         'loss of $32k. Support is the only department covering its allocated '
         'overheads. The revenue variance arose almost entirely in Integration '
         'at a 10.1% gross margin, contributing turnover with very little '
         'profit. At the year-to-date gross margin of 38.9%, monthly revenue of '
         'approximately $549k is required to cover overheads, against a current '
         'average of $541k.')
tf = by['Text 53'].text_frame
runs = tf.paragraphs[0].runs
runs[0].text = 'The month in summary. '
runs[1].text = STRIP
for r in runs[2:]:
    r._r.getparent().remove(r._r)

p.save(OUT)
print('slide 3 rebuilt with the company total card')
