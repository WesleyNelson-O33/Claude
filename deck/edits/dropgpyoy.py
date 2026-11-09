"""Page 27 comes out, and its one useful finding moves onto page 26.

The gross profit year on year page was the only page in the deck that needed a
disclaimer to stop it being read as the net result. A page that needs a caveat is
doing a poor job, so it goes.

What was worth keeping off it was Integration: the biggest rise in gross profit of
the five, earned on revenue eleven times last September, which at a 10% margin left
only $10,471 after overhead. That now sits on page 26 where the reader already has
the result in front of them.

Run in two passes. python-pptx hands two parts the same name if a slide is added
and another removed in the same pass, so the commentary is rewritten first and the
page removed second.
"""
import sys
sys.path.insert(0, '/home/user/Claude/deck/edits')
from pptx import Presentation
from addslide import relay

SNIP = '/home/user/Claude/deck/snips'

BULLETS = [
    'Two of the five made a profit. Support kept $29k and Integration $10k.',
    'Events lost $54k, Consulting $20k and Video $7k. Together the five are the '
    '$41k loss.',
    'All five made a gross profit. It is the overhead charged to them that puts '
    'three below the line.',
    'Integration billed eleven times last September, but a 10% margin left only '
    '$10k after overhead.',
]
NOTES = """The result by department for September 2026 after the overhead charged to each one, taken from section 7 of the Overhead Split tab. Teal is above the line, red is below.

Support $29,007. Integration $10,471. Video a loss of $6,846. Consulting a loss of $19,870. Events a loss of $54,227. The five add to the $41,464 loss for the month.

Two of the five made a profit.

Every department made a gross profit. Support $101,290, Events $56,881, Integration $23,756, Video $5,813, and Consulting $74. On gross profit nobody lost money. What puts three of them below the line is the overhead charged to them: $229,278 across the month against $187,814 of gross profit earned. The company cannot cover its overhead base out of the margin it is making.

Events is the clearest case. It carries $111,108 of overhead, 48.5% of the company total, on 13.0% of the revenue. Its margin is the second best of the five at 77.6% and it still posts the largest loss on the page. Improving margin does not fix that. Either Events carries more revenue or it carries less overhead.

Consulting is a different problem. It earned $74 of gross profit on $3,531 of revenue and carries $19,944 of overhead, so almost the whole of its loss is overhead on a department that is barely trading.

Integration is the one to watch and it is easy to miss. Its revenue went from $21,421 last September to $234,858 this September, close to eleven times. Its gross profit rose from $5,363 to $23,756, the largest increase of the five. But that revenue carries a 10.1% margin against Support's 41.7%, so after $13,285 of overhead only $10,471 reached the line. Integration is now running the second largest revenue in the business for the smallest return on it.

Video's loss of $6,846 is the smallest and the least structural. It billed $7,303 against $20,025 in the same month last year, so it is a volume problem on a department with a 79.6% margin.

Two points for the board. Improving margin does not fix Events or Consulting, because their losses are overhead rather than margin. And the overhead keys were set for FY27 and have not been tested since, which is the action already on the recommendations page."""

p = Presentation(sys.argv[1])
w, h, strip = relay(p.slides[25], f'{SNIP}/dept_line.png', BULLETS, NOTES,
                    'snip dept_line')
print(f'page 26 commentary updated | picture {w:.2f} x {h:.2f}, strip {strip:.2f}')
p.save(sys.argv[2])
