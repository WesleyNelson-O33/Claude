"""Headings reworded, and everything put on the same grid.

Three jobs in one pass.

The profit page comes apart into a monthly page and a quarterly page, because the
deck serves both reviews and either page may be hidden.

The headings are rewritten into the register a board paper uses. Plain nouns,
the month written out, and the section names, the eyebrows and the agenda lines
all saying the same thing as each other.

Then the grid. The content pages were already square: the commentary panel sits
on the left margin, runs the full content width and its bottom edge lands on 5.06
on every page. What was out was the title box, which carried five different
depths, and the covers and section dividers, where the section name and the
wordmark had drifted by a tenth or two.
"""
import copy, sys
sys.path.insert(0, '/home/user/Claude/deck/edits')
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from addslide import (add, relay, balance, need, write,
                      strip_depth, FLOOR, LEFT, SIZE)
from profitsplit import MONTH_B, MONTH_N, QTR_B, QTR_N

SNIP = '/home/user/Claude/deck/snips'

# ---- the headings ----------------------------------------------------------
# slide number in the deck AFTER the profit split -> new title
TITLES = {
    2:  'Scope of this review',
    3:  'Performance summary, September 2026',
    5:  'Profit and loss summary, income',
    6:  'Monthly revenue by department',
    7:  'Revenue against budget and prior years',
    8:  'Quarterly revenue, three year comparison',
    9:  'Monthly revenue, three year history',
    10: 'Year to date revenue against budget',
    11: 'Profit and loss summary, cost and profit',
    12: 'Monthly expenditure against budget',
    13: 'Quarterly expenditure against budget',
    14: 'Monthly net result, three year comparison',
    15: 'Quarterly net result, three year comparison',
    16: 'Departmental result, September 2026',
    17: 'Results by reporting group, September 2026',
    18: 'Support, September 2026',
    19: 'Events, September 2026',
    20: 'Video, September 2026',
    21: 'Integration, September 2026',
    22: 'Consulting, September 2026',
    23: 'Departmental revenue against budget and prior year',
    25: 'Operating cost distribution, year to date',
    26: 'Other operating expenses, year to date',
    27: 'Direct and indirect labour, year to date',
    28: 'Employment expense by function',
    30: 'Top ten customers by revenue, year to date',
    31: 'Debtor days and collection performance',
    32: 'Aged receivables by customer',
    34: 'Chargeable percentage by department, September 2026',
    35: 'Hours mix by department, September 2026',
    36: 'Hours mix by reporting group, September 2026',
    38: 'Full year revenue against budget',
    39: 'Recommendations and agreed actions',
}
EYEBROWS = {
    'INTRODUCTION': 'EXECUTIVE SUMMARY',
    'WHAT WE SPEND': 'OPERATING EXPENDITURE',
    'CUSTOMERS AND CASH': 'CUSTOMERS AND RECEIVABLES',
    'SUMMARY AND ACTIONS': 'RECOMMENDATIONS',
}
SECTIONS = {                       # divider headline and agenda line
    'What we spend': 'Operating expenditure',
    'Customers and cash': 'Customers and receivables',
    'Links': 'Source documents',
}

# ---- the grid --------------------------------------------------------------
TITLE_TOP, TITLE_H = 0.70, 0.40
DIV_LEFT, DIV_HEAD_TOP, DIV_SUB_TOP, DIV_SUB_STEP = 1.75, 2.05, 2.62, 0.44
MARK_SIZE, MARK_TOP = 40.0, (0.62, 1.05)


def set_text(shape, text):
    para = shape.text_frame.paragraphs[0]
    if para.runs:
        para.runs[0].text = text
        for r in list(para.runs)[1:]:
            r._r.getparent().remove(r._r)


p = Presentation(sys.argv[1])

# 1 -- the profit page comes apart
slide = p.slides[13]
relay(slide, f'{SNIP}/profit_month.png', MONTH_B, MONTH_N,
      'snip profit_line', 'Monthly net result, three year comparison')
for sh in slide.shapes:
    if sh.name == 'snip profit_line':
        sh.name = 'snip profit_month'
add(p, 8, 14, 'FINANCIAL PERFORMANCE', 'Quarterly net result, three year comparison',
    f'{SNIP}/profit_qtr.png', QTR_B, QTR_N, 'snip profit_qtr')
print('profit page split into a monthly page and a quarterly page')

# 2 -- headings, eyebrows, section names and the agenda lines
for i, s in enumerate(p.slides, start=1):
    for sh in s.shapes:
        if not sh.has_text_frame:
            continue
        txt = sh.text_frame.text.strip()
        if sh.name == 'Text 1' and i in TITLES:
            set_text(sh, TITLES[i])
        elif sh.name == 'Text 0' and txt in EYEBROWS:
            set_text(sh, EYEBROWS[txt])
        elif txt in SECTIONS:
            set_text(sh, SECTIONS[txt])
print(f'{len(TITLES)} titles reworded, {len(EYEBROWS)} eyebrows, '
      f'{len(SECTIONS)} section names')

# 3 -- the grid
for i, s in enumerate(p.slides, start=1):
    title = next((sh for sh in s.shapes if sh.name == 'Text 1'), None)
    if title is not None:
        title.top, title.height = Inches(TITLE_TOP), Inches(TITLE_H)
        for para in title.text_frame.paragraphs:
            para.alignment = PP_ALIGN.CENTER

    # section dividers: the numbered square, the headline and the strapline
    if any(sh.name == 'Shape 1' for sh in s.shapes):
        head = next((sh for sh in s.shapes if sh.name == 'Text 4'), None)
        if head is not None:
            head.left, head.top = Inches(DIV_LEFT), Inches(DIV_HEAD_TOP)
        subs = sorted((sh for sh in s.shapes if sh.name == 'Text 5'),
                      key=lambda sh: sh.top)
        for k, sub in enumerate(subs):
            sub.left = Inches(DIV_LEFT)
            sub.top = Inches(DIV_SUB_TOP + k * DIV_SUB_STEP)

    # the two cover pages: the wordmark and the line under it
    mark = next((sh for sh in s.shapes
                 if sh.name == 'Text 3' and sh.text_frame.text.strip() == 'cts'), None)
    if mark is not None:
        for para in mark.text_frame.paragraphs:
            for r in para.runs:
                r.font.size = Pt(MARK_SIZE)
        sub = next((sh for sh in s.shapes if sh.name == 'Text 4'), None)
        if sub is not None:
            sub.left, sub.top = Inches(MARK_TOP[0]), Inches(MARK_TOP[1])
print('title boxes, section dividers and cover wordmarks squared up')

# 4 -- every commentary strip onto the same grid.
#
# Twenty five pages name the second column 'Text 5b' and four name it 'Text 6'.
# Looking only for 'Text 5b' put a new box on top of the four that use the other
# name, so the second column is now found under either name and renamed.
#
# The bullets are gathered from both columns and dealt out again, so the split
# falls in the same place on every page, and the strip is snapped up to the
# nearest rung on the ladder so the deck has three strip depths rather than six.
done = []
for i, s in enumerate(p.slides, start=1):
    panel = next((sh for sh in s.shapes if sh.name == 'commentary panel'), None)
    if panel is None or panel.top / 914400 < 3.0:
        continue
    body = next((sh for sh in s.shapes if sh.name == 'Text 5'), None)
    if body is None:
        continue
    extra = next((sh for sh in s.shapes if sh.name in ('Text 5b', 'Text 6')), None)
    if extra is None:
        tree = s.shapes._spTree
        el = copy.deepcopy(body._element)
        tree.append(el)
        extra = next(sh for sh in s.shapes if sh._element is el)
    extra.name = 'Text 5b'

    bullets = [x.text for x in body.text_frame.paragraphs if x.text.strip()]
    bullets += [x.text for x in extra.text_frame.paragraphs if x.text.strip()]
    if len(bullets) < 2:
        continue
    cols = list(balance(bullets, 4.05))
    ph = strip_depth(max(need(c, 4.05) for c in cols if c) + 0.30)
    strip_top = FLOOR - ph

    panel.left, panel.top = Inches(LEFT), Inches(strip_top)
    panel.width, panel.height = Inches(8.90), Inches(ph)
    label = next((sh for sh in s.shapes if sh.name == 'Text 4'), None)
    if label is not None:
        label.left, label.top = Inches(0.75), Inches(strip_top + 0.08)
        label.width, label.height = Inches(3.0), Inches(0.25)
    for box, x, items in ((body, 0.75, cols[0]), (extra, 5.10, cols[1])):
        box.left, box.top = Inches(x), Inches(strip_top + 0.27)
        box.width, box.height = Inches(4.05), Inches(ph - 0.30)
        write(box, items)
        for para in box.text_frame.paragraphs:
            para.alignment = None
            for r in para.runs:
                r.font.size = Pt(SIZE)
    done.append(i)
print(f'{len(done)} commentary strips put on the grid')

# 5 -- the actions table has a column of dates headed "Suggested by", which
# reads as a name rather than a deadline.
for s_ in p.slides:
    for sh in s_.shapes:
        if not getattr(sh, 'has_table', False):
            continue
        for cell in sh.table.rows[0].cells:
            if cell.text.strip() == 'Suggested by':
                para = cell.text_frame.paragraphs[0]
                para.runs[0].text = 'Target date'
                for r in list(para.runs)[1:]:
                    r._r.getparent().remove(r._r)
                print('actions table: "Suggested by" -> "Target date"')

p.save(sys.argv[2])
