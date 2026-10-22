"""The departmental profit and loss chart, after the last department page.

Page 18 keeps gross profit against overhead but drops the diamond: the result is
a third bar now, and the labels sit out at the bar ends. It is retitled so it and
the new page are not both called the departmental result.

The new page carries the whole profit and loss by department. What each one
earned, what it cost, and what was left. Total cost is cost of sales plus
overhead, which is why a figure like Support's 213,942 is on no row of the
printed table: it is 141,659 of cost of sales plus 72,283 of overhead.

Both are monthly pages.
"""
import sys
sys.path.insert(0, '/home/user/Claude/deck/edits')
from pptx import Presentation
from addslide import add

SNIP = '/home/user/Claude/deck/snips'

BULLETS = [
    'Support earned $243k, cost $214k and kept $29k.',
    'Events earned $73k, cost $128k and lost $54k.',
    'Integration earned $235k, cost $224k and kept $10k, the most work for the '
    'least return.',
    'Across all five, $562k earned against $603k of cost. That is the $41k loss.',
]
NOTES = """The whole profit and loss by department for September 2026, taken from section 7 of the Overhead Split tab. Teal is what the department earned, red is what it cost, navy is what was left.

Support earned $242,948, cost $213,942 and kept $29,007.
Events earned $73,304, cost $127,531 and lost $54,227.
Video earned $7,303, cost $14,149 and lost $6,846.
Integration earned $234,858, cost $224,387 and kept $10,471.
Consulting earned $3,531, cost $23,401 and lost $19,870.

Across the five, $561,945 of revenue against $603,408 of cost, which is the $41,464 loss for the month.

On the cost figure, because it does not appear on the department pages. Total cost is cost of sales plus overhead. For Support that is $141,659 of cost of sales plus $72,283 of overhead, giving $213,942. Cost of sales is not a line on those tables either, it is revenue less gross profit. Every department on this chart ties to the cent: revenue less total cost equals the result shown on the table, and the five results add to the $41,464 net loss on the P&L.

What the chart shows that the department pages cannot. Support and Integration have almost the same revenue, $242,948 and $234,858, and almost the same cost, $213,942 and $224,387. The results are $29,007 and $10,471. Integration is running nearly the same volume for a third of the return, because its gross margin is 10.1% against Support's 41.7%.

Events is the opposite problem. It is the smallest cost base of the three large departments and the second largest loss, because $111,108 of the $127,531 it costs is allocated overhead rather than cost of sales.

Video and Consulting are small on this chart because they are small. Between them they earned $10,834 and cost $37,550."""

p = Presentation(sys.argv[1])

title = next(sh for sh in p.slides[17].shapes if sh.name == 'Text 1')
para = title.text_frame.paragraphs[0]
para.runs[0].text = 'Gross profit against overhead, September 2026'
for r in list(para.runs)[1:]:
    r._r.getparent().remove(r._r)
print('page 18 retitled')

w, h, strip = add(p, 8, 24, 'DEPARTMENTAL PERFORMANCE',
                  'Departmental profit and loss, September 2026',
                  f'{SNIP}/dept_pl.png', BULLETS, NOTES, 'snip dept_pl')
print(f'page 25: the departmental P&L | picture {w:.2f} x {h:.2f}, strip {strip:.2f}')

p.save(sys.argv[2])
