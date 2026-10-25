"""Push the new overhead allocation through the deck.

The Staff key now matches the P&L analysis, so every departmental result moves.
Revenue, cost of sales and gross profit are untouched, and so is the company
result. What changes is how the overhead is divided and therefore each
department's bottom line.

September, before and after:
  Support      29,007 -> 21,762      overhead 72,283 -> 79,527
  Events      -54,227 -> -46,265     overhead 111,108 -> 103,146
  Video        -6,846 -> -6,617      overhead 12,659 -> 12,430
  Integration  10,471 -> 10,548      overhead 13,285 -> 13,208
  Consulting  -19,870 -> -20,892     overhead 19,944 -> 20,966

Six pictures go back from file. The five department pages keep their crop: their
snip is an eighteen row table with the seven overhead detail rows cut out, so the
fresh snip is cropped the same way before it goes back on.
"""
import io, sys
sys.path.insert(0, '/home/user/Claude/deck/edits')
from PIL import Image
from pptx import Presentation
from pptx.util import Inches
from relayout import drop_rows
from addslide import relay

SNIP = '/home/user/Claude/deck/snips'
WORK = '/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a/scratchpad/recrop'
OVERHEAD_DETAIL = {9, 10, 11, 12, 13, 14, 15}

PLAIN = {17: 'dept_result', 18: 'dept_mix', 19: 'dept_three',
         25: 'dept_pl', 26: 'dept_line'}
CROPPED = {20: 'dept_Support', 21: 'dept_Events', 22: 'dept_Video',
           23: 'dept_Integration', 24: 'dept_Consulting'}

# page 3, the per department cards: the after overhead line only
CARDS = {'+$29k after overhead': '+$22k after overhead',
         '$(54)k after overhead': '$(46)k after overhead',
         '+$10k after overhead': '+$11k after overhead',
         '$(20)k after overhead': '$(21)k after overhead'}

MIX_B = [
    'Gross profit of $188k did not cover overheads of $229k. That gap is the '
    '$41k loss.',
    'Only Support and Integration covered the overhead charged to them.',
    'Events carries 45% of all overhead on 13% of revenue, so a 78% margin still '
    'ends in a $46k loss.',
    'Consulting earned $74 of gross profit on $4k of revenue and carries $21k of '
    'overhead.',
]
MIX_N = """Gross profit against the overhead charged to each department for September 2026, taken from section 7 of the Overhead Split tab. The bar above the line is what the department earned, the bar below is the overhead allocated to it, and the third bar is the result.

Support earned $101,290 and carries $79,527, leaving $21,762.
Events earned $56,881 and carries $103,146, leaving a loss of $46,265.
Video earned $5,813 and carries $12,430, leaving a loss of $6,617.
Integration earned $23,756 and carries $13,208, leaving $10,548.
Consulting earned $74 and carries $20,966, leaving a loss of $20,892.

The company earned $187,814 of gross profit against $229,278 of overhead. The $41,464 difference is the loss for the month. The business did not earn enough gross profit to carry its own overhead base.

Events is still the clearest case. It carries 45.0% of all overhead on 13.0% of the revenue, against a gross margin of 77.6%, the second highest of the five. No improvement in margin fixes that.

Integration is the other way round. It carries 5.8% of the overhead on 41.8% of the revenue, so it covers its allocation easily and still only returns $10,548, because its gross margin is 10.1%.

Support is the one department behaving as it should. A 41.7% margin on contract revenue, covering its allocated overhead, and returning $21,762.

The allocation behind these figures now matches the P&L analysis. The Staff key is driven by a headcount of 49 Support, 18 Events, 3 Video, 1 Integration and 1 Consulting."""

THREE_B = [
    'Support is the only group above the line, at $22k on $243k of revenue.',
    'Production lost $53k. A 78% margin on $81k of revenue cannot carry $116k of '
    'overhead.',
    'Consulting lost $10k on $238k of revenue, because its margin is 10%.',
    'The three groups are the $41k loss for the month.',
]
THREE_N = """The same result rolled up the way the P&L groups the business, taken from section 8 of the Overhead Split tab. Production is Events and Video. Consulting is Integration and Consulting.

Support: revenue $242,948, gross profit $101,290 at 41.7%, overhead $79,527, result $21,762.
Production: revenue $80,607, gross profit $62,694 at 77.8%, overhead $115,577, result a loss of $52,882.
Consulting: revenue $238,389, gross profit $23,830 at 10.0%, overhead $34,174, result a loss of $10,344.

The three add to the $41,464 loss for the month.

Set out this way the business has one group covering its costs and two that do not, and they fail for opposite reasons.

Production earns a 77.8% margin, the best in the business, on $80,607 of revenue. That produces $62,694 of gross profit against $115,577 of overhead. The margin is not the problem. The revenue is too small for the cost base sitting on it.

Consulting is the mirror image. It brought in $238,389, almost as much as Support, and converted it into $23,830 at a 10.0% margin. Its overhead is only $34,174 and it still loses $10,344, because there is almost no margin on the work.

Support carries the business: 41.7% on $242,948, covering $79,527 of overhead and returning $21,762.

The allocation behind these figures now matches the P&L analysis."""

PL_B = [
    'Support earned $243k, cost $221k and kept $22k.',
    'Events earned $73k, cost $120k and lost $46k.',
    'Integration earned $235k, cost $224k and kept $11k, the most work for the '
    'least return.',
    'Across all five, $562k earned against $603k of cost. That is the $41k loss.',
]
PL_N = """The whole profit and loss by department for September 2026, taken from section 7 of the Overhead Split tab. Teal is what the department earned, red is what it cost, navy is what was left.

Support earned $242,948, cost $221,186 and kept $21,762.
Events earned $73,304, cost $119,569 and lost $46,265.
Video earned $7,303, cost $13,920 and lost $6,617.
Integration earned $234,858, cost $224,311 and kept $10,548.
Consulting earned $3,531, cost $24,423 and lost $20,892.

Across the five, $561,945 of revenue against $603,408 of cost, which is the $41,464 loss for the month.

Total cost is cost of sales plus the overhead charged to the department. Cost of sales is not a line on the department pages either, it is revenue less gross profit.

Support and Integration have almost the same revenue, $242,948 and $234,858, and almost the same cost, $221,186 and $224,311. The results are $21,762 and $10,548. Integration is running nearly the same volume for half the return, because its gross margin is 10.1% against Support's 41.7%.

Events is the opposite problem. It is the smallest cost base of the three large departments and the second largest loss, because $103,146 of the $119,569 it costs is allocated overhead rather than cost of sales.

The allocation behind these figures now matches the P&L analysis."""

LINE_B = [
    'Two of the five made a profit. Support kept $22k and Integration $11k.',
    'Events lost $46k, Consulting $21k and Video $7k. Together the five are the '
    '$41k loss.',
    'All five made a gross profit. It is the overhead charged to them that puts '
    'three below the line.',
    'Integration billed eleven times last September, but a 10% margin left only '
    '$11k after overhead.',
]
LINE_N = """The result by department for September 2026 after the overhead charged to each one, taken from section 7 of the Overhead Split tab. Teal is above the line, red is below.

Support $21,762. Integration $10,548. Video a loss of $6,617. Consulting a loss of $20,892. Events a loss of $46,265. The five add to the $41,464 loss for the month.

Two of the five made a profit.

Every department made a gross profit. Support $101,290, Events $56,881, Integration $23,756, Video $5,813, and Consulting $74. On gross profit nobody lost money. What puts three of them below the line is the overhead charged to them: $229,278 across the month against $187,814 of gross profit earned.

Events is the clearest case. It carries 45.0% of all overhead on 13.0% of the revenue. Its margin is the second best of the five at 77.6% and it still posts the largest loss on the page. Improving margin does not fix that. Either Events carries more revenue or it carries less overhead.

Consulting is a different problem. It earned $74 of gross profit on $3,531 of revenue and carries $20,966 of overhead, so almost the whole of its loss is overhead on a department that is barely trading.

Integration is the one to watch and it is easy to miss. Its revenue went from $21,421 last September to $234,858 this September, close to eleven times. But that revenue carries a 10.1% margin against Support's 41.7%, so after $13,208 of overhead only $10,548 reached the line.

Video's loss of $6,617 is the smallest and the least structural. It billed $7,303 against $20,025 in the same month last year, so it is a volume problem on a department with a 79.6% margin.

The allocation behind these figures now matches the P&L analysis. The Staff key is driven by a headcount of 49 Support, 18 Events, 3 Video, 1 Integration and 1 Consulting, which is the basis the analysis uses."""


def swap(slide, path, name):
    old = next(sh for sh in slide.shapes if sh.name == name)
    geom = (old.left, old.top, old.width, old.height)
    tree = slide.shapes._spTree
    tree.remove(old._element)
    pic = slide.shapes.add_picture(path, geom[0], geom[1],
                                   width=geom[2], height=geom[3])
    pic.name = name
    tree.remove(pic._element)
    tree.insert(2, pic._element)


import os
os.makedirs(WORK, exist_ok=True)
p = Presentation(sys.argv[1])

for page, tag in PLAIN.items():
    swap(p.slides[page - 1], f'{SNIP}/{tag}.png', f'snip {tag}')
print(f'pictures replaced on pages {sorted(PLAIN)}')

for page, tag in CROPPED.items():
    im = Image.open(f'{SNIP}/{tag}.png').convert('RGB')
    out = drop_rows(im, 18, OVERHEAD_DETAIL)
    path = f'{WORK}/{tag}.png'
    out.save(path)
    swap(p.slides[page - 1], path, f'snip {tag}')
print(f'department pages re-cropped and replaced: {sorted(CROPPED)}')

card = p.slides[2]
hits = 0
for sh in card.shapes:
    if not sh.has_text_frame:
        continue
    t = sh.text_frame.text.strip()
    if t in CARDS:
        para = sh.text_frame.paragraphs[0]
        para.runs[0].text = CARDS[t]
        for r in list(para.runs)[1:]:
            r._r.getparent().remove(r._r)
        hits += 1
print(f'page 3 cards updated: {hits}')

for page, snip_tag, bullets, notes in (
        (18, 'dept_mix', MIX_B, MIX_N), (19, 'dept_three', THREE_B, THREE_N),
        (25, 'dept_pl', PL_B, PL_N), (26, 'dept_line', LINE_B, LINE_N)):
    relay(p.slides[page - 1], f'{SNIP}/{snip_tag}.png', bullets, notes,
          f'snip {snip_tag}')
    print(f'page {page} commentary rewritten')

p.save(sys.argv[2])
