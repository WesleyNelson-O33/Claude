"""Gross profit by department, this September against last.

What was asked for was profit and loss by department year on year, and only part
of that exists. Revenue and cost of sales are held by department on both year
tabs, so gross profit compares. Overhead was not split by department before this
financial year, which is why the department pages print "not split last year"
against Total Expenses and Net Profit/Loss, so the net result cannot go back.
September 2024 is not held by department anywhere in the pack.

The page says so rather than leaving the reader to assume this is the net result.

This is a monthly page.
"""
import sys
sys.path.insert(0, '/home/user/Claude/deck/edits')
from pptx import Presentation
from addslide import add

SNIP = '/home/user/Claude/deck/snips'

BULLETS = [
    'Company gross profit was $188k against $162k last September, up $26k.',
    'Integration is up $18k, Events $9k and Support $7k.',
    'Video fell from $13k to $6k and Consulting from $2k to almost nothing.',
    'Gross profit only. Overhead was not split by department before this year, '
    'so the net result cannot go back.',
]
NOTES = """Gross profit by department for September 2026 against September 2025, taken from the Actual FY26 and P&L FY27 tabs.

Support $101,290 against $94,701, up $6,589.
Events $56,881 against $47,493, up $9,388.
Video $5,813 against $12,625, down $6,812.
Integration $23,756 against $5,363, up $18,394.
Consulting $74 against $2,085, down $2,011.
Company $187,814 against $162,267, up $25,548.

Read it with the revenue beside it, because the two do not move together.

Integration added $18,394 of gross profit, the largest increase of the five, and did it on revenue that rose from roughly $56k to $234,858. So the extra gross profit came at about a 10% margin, which is why Integration appears strong here and only returns $10,471 after overhead.

Events added $9,388 of gross profit on revenue that barely moved, which is a genuine margin improvement rather than a volume one.

Video went the other way. Gross profit more than halved because the department billed $7,303 against $20,025 last September. That is a volume problem, not a margin one: its margin is still 79.6%.

Now the limitation, and it matters for how this page is read.

This is gross profit, not the net result. Overhead was not allocated to departments before this financial year. The department pages in this deck print "not split last year" against Total Expenses and Net Profit/Loss for exactly that reason, and the pack has no department split of overhead for FY26 to work from. So the net result by department, which is the figure on the previous page, cannot be shown against last year.

September 2024 by department cannot be shown either. The pack holds departmental detail for FY26 and FY27 only, and Xero's profit and loss is by account rather than by department, so it cannot be rebuilt reliably from there.

If the board wants a three year net result by department, the overhead keys would have to be applied backwards to FY26 and FY25. That is a decision about method, not a calculation, and I have not done it."""

p = Presentation(sys.argv[1])
w, h, strip = add(p, 8, 25, 'DEPARTMENTAL PERFORMANCE',
                  'Gross profit by department, year on year',
                  f'{SNIP}/dept_gp_yoy.png', BULLETS, NOTES, 'snip dept_gp_yoy')
print(f'page 26: gross profit year on year | picture {w:.2f} x {h:.2f}, strip {strip:.2f}')
p.save(sys.argv[2])
