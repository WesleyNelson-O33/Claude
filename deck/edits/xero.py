"""Three years of revenue history, now that it has been pulled from Xero.

Pages 8 and 9 are rebuilt because the comparison they show has changed: FY25 is
no longer an empty column waiting for a paste, it is a Xero pull, so both charts
now carry three years rather than two. A third page is added for the full thirty
six month run.

The pages are rebuilt rather than patched because the depth of the commentary
strip is set by how much commentary there is, and the commentary has changed.
"""
import sys
sys.path.insert(0, '/home/user/Claude/deck/edits')
from pptx import Presentation
from addslide import add, balance, need, write, TOP, FLOOR, LEFT, SIZE
from PIL import Image
from pptx.util import Inches, Pt

SNIP = '/home/user/Claude/deck/snips'

A_BULLETS = [
    'September revenue was $562k, against $356k last September and $440k the '
    'September before that.',
    'The three year shape is a dip and a recovery. September fell 19% between '
    'FY25 and FY26, then rose 58% into FY27.',
    'Measured against FY25 rather than the weak FY26, September is 28% ahead, '
    'which is the steadier read on growth.',
    'Xero holds the last three financial years, so September 2024 is the '
    'earliest month available for this comparison.',
]
A_NOTES = """Revenue for September across three years, taken from the Xero History tab and the Cost Analysis tab.

September 2026 $561,945. September 2025 $356,323. September 2024 $440,243.

September 2026 is $205,621 or 57.7% ahead of September 2025, and $121,701 or 27.6% ahead of September 2024. September 2025 was itself $83,920 or 19.1% below September 2024.

The 57.7% headline is flattering because the year it is measured against was the weakest of the three. FY26 was a down year across the board, and part of what FY27 has done is recover ground rather than break new ground. Against FY25 the same month is 27.6% ahead, which over two years is about 13% a year, and that is the number worth carrying into the reforecast.

On the source. FY25 was pulled from Xero on 22 October 2026 as total income by month on the accrual basis, and sits on its own Xero History tab in the pack. FY26 and FY27 are references to the pack's own actuals tabs. Three months were checked against Xero directly and agree: July 2025, August 2025 and September 2025 on the FY26 side, and September 2026 on the FY27 side.

A five year comparison was asked for and cannot be produced. Xero only serves the last three financial years, and a request for September 2022 is refused by the API. FY23 and FY24 would have to come from an older copy of the pack or from a saved export. The columns for them are in place on the Cost Analysis tab and the chart reads whichever years carry figures."""

B_BULLETS = [
    'First quarter revenue was $1,624k, against $1,229k last year and $1,380k '
    'the year before.',
    'FY26 was the softest of the three, 11% below FY25 for the quarter and 13% '
    'below for the full year.',
    'Against FY26 the quarter is 32% ahead. Against FY25 it is 18% ahead, which '
    'is about 8.5% a year.',
    'The growth is on revenue alone. The quarter delivered a 38.9% margin and '
    '$5.8k of EBITDA.',
]
B_NOTES = """First quarter revenue across three years, taken from the Xero History tab and the Cost Analysis tab.

Q1 2026, being July to September 2026, $1,624,460. Q1 2025 $1,228,582. Q1 2024 $1,380,102.

Against last year the quarter is $395,878 or 32.2% ahead. Against two years ago it is $244,358 or 17.7% ahead. The middle year fell $151,520 or 11.0% below the first.

The two readings tell different stories and the second is the more useful one. A 32% increase measured off the weakest of three years is not a growth rate the business can plan on. Compounded over the two years, the first quarter has grown about 8.5% a year. That is the figure to put against the reforecast.

The first quarter has held a steady share of the year in both completed years, 19.7% of FY25 and 20.1% of FY26. On that relationship a first quarter of $1,624,460 implies a full year of around $8.1 million, which is well ahead of both completed years and ahead of where the FY27 budget sits after the first three months.

What the quarter did not do is reach the result. The margin delivered on that revenue was 38.9% and the quarter produced $5,774 of EBITDA. Revenue growth at this margin does not reach the bottom line, and that is the point the rest of the pack returns to.

FY23 and FY24 are outside the three financial years Xero serves, so the comparison stops at FY25."""

C_BULLETS = [
    'Thirty six months of revenue, from July 2024 through to September 2026.',
    'FY26 was the softer year at $6,113k against $7,000k, and it was down in '
    'both halves.',
    'The largest single swing is January: $918k in FY25 against $269k in FY26.',
    'October carries the year. It was $949k in FY25 and $1,036k in FY26, and '
    'FY27 has it budgeted at $950k.',
]
C_NOTES = """Thirty six months of revenue, taken from the new Xero History tab.

FY25 totalled $7,000,357 and FY26 $6,113,186, a fall of $887,171 or 12.7%. FY27 stands at $1,624,460 after three months.

The fall in FY26 was not confined to one half. The first half was $533,509 or 14.6% below FY25 and the second half $353,662 or 10.6% below.

Two features of the shape are worth the board's attention.

The first is October. It was the largest month of FY25 at $948,991 and the largest month of FY26 at $1,035,562, and the FY27 budget repeats that pattern with $950,416. October is a quarter of the way to a seven figure month on its own, and if it behaves like the rest of FY27 rather than like its own history, the full year gap against budget widens materially. There is no second October in the year to recover it.

The second is January. FY25 billed $918,289 in January and FY26 billed $268,631, a swing of $649,658 in a single month. January is the month most exposed to how much work carries over the Christmas shutdown, and it is the month least able to be recovered once it is gone.

On the source and how it is maintained. FY25 was pulled from Xero on 22 October 2026 as total income by month on the accrual basis, and is held on the tab as figures, because the pack has no FY25 tab to reference. FY26 and FY27 are references to the Actual FY26 and P&L FY27 tabs, so they follow the pack rather than duplicating it, and FY27 shows closed months only. Four months were checked against Xero directly and agree to the cent, apart from one cent of rounding on September 2025.

Xero serves the last three financial years. FY23 and FY24 cannot be pulled, which is why the history begins in July 2024."""


def relay(slide, snip, bullets, notes, name, title=None):
    """Rework a chart page that was built by add(): new picture, new commentary,
    and the strip resized to whatever depth the new commentary needs. The page
    itself is left alone, so nothing is deleted from the deck."""
    by = {sh.name: sh for sh in slide.shapes}
    tree = slide.shapes._spTree
    if title is not None:
        by['Text 1'].text_frame.paragraphs[0].runs[0].text = title

    cols = list(balance(bullets, 4.05))
    ph_panel = max(need(c, 4.05) for c in cols if c) + 0.30
    strip_top = FLOOR - ph_panel

    old = by[name]
    tree.remove(old._element)
    im = Image.open(snip)
    aspect = im.width / im.height
    pic_h = min(strip_top - 0.12 - TOP, 8.90 / aspect)
    pic_w = pic_h * aspect
    pic = slide.shapes.add_picture(snip, Inches(LEFT + (8.90 - pic_w) / 2),
                                   Inches(TOP), width=Inches(pic_w),
                                   height=Inches(pic_h))
    pic.name = name
    tree.remove(pic._element)
    tree.insert(2, pic._element)

    panel, label = by['commentary panel'], by['Text 4']
    panel.left, panel.top = Inches(LEFT), Inches(strip_top)
    panel.width, panel.height = Inches(8.90), Inches(ph_panel)
    label.left, label.top, label.width = Inches(0.75), Inches(strip_top + 0.08), Inches(3.0)
    for key, x, items in (('Text 5', 0.75, cols[0]), ('Text 5b', 5.10, cols[1])):
        box = by[key]
        box.left, box.top = Inches(x), Inches(strip_top + 0.27)
        box.width, box.height = Inches(4.05), Inches(ph_panel - 0.30)
        write(box, items)
        for para in box.text_frame.paragraphs:
            for r in para.runs:
                r.font.size = Pt(SIZE)

    slide.notes_slide.notes_text_frame.text = notes
    return pic_w, pic_h, ph_panel


p = Presentation(sys.argv[1])

for idx, title, snip, bullets, notes, name in (
    (7, 'Revenue for the month, year on year',
     f'{SNIP}/rev_years.png', A_BULLETS, A_NOTES, 'snip rev_years'),
    (8, 'First quarter revenue, year on year',
     f'{SNIP}/rev_q1.png', B_BULLETS, B_NOTES, 'snip rev_q1'),
):
    w, h, strip = relay(p.slides[idx], snip, bullets, notes, name, title)
    print(f'page {idx + 1}: {title} | picture {w:.2f} x {h:.2f}, strip {strip:.2f}')

w, h, strip = add(p, 6, 9, 'FINANCIAL PERFORMANCE', 'Three years of revenue by month',
                  f'{SNIP}/xero_hist.png', C_BULLETS, C_NOTES, 'snip xero_hist')
print(f'page 10: Three years of revenue by month | picture {w:.2f} x {h:.2f}, strip {strip:.2f}')

p.save(sys.argv[2])
