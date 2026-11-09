"""Two slides: revenue by month year on year, and the first quarter year on
year. Both state plainly that only two years are in the pack."""
import sys
sys.path.insert(0, '/home/user/Claude/deck/edits')
from pptx import Presentation
from addslide import add

A_BULLETS = [
    'FY27 is ahead of FY26 in all three closed months: $459k against $321k, '
    '$603k against $551k, $562k against $356k.',
    'The quarter is $1,624k against $1,229k, up 32%.',
    "FY26's own shape is the warning. Its largest month was October at $1,036k, "
    'and the FY27 budget repeats that pattern.',
    'Only FY26 and FY27 are in the pack. FY23, FY24 and FY25 need one Xero '
    'export each.',
    'The chart reads whichever years have figures, so pasting a year in is all '
    'that is needed.',
]
A_NOTES = """Revenue by month, FY27 against FY26, taken from the Cost Analysis tab.

FY27: July $459,049, August $603,467, September $561,945.
FY26: July $320,808, August $551,451, September $356,323.

FY27 is ahead in each of the three closed months, and ahead by $395,878 or 32.2% for the quarter.

The FY27 line stops at September because the months beyond it hold budget rather than actuals, and charting those would show a plan as though it were a result.

What FY26's full year shows is worth noting. Its largest month was October at $1,035,562, and its second half ran well below that. The FY27 budget repeats the same shape, with October budgeted at $950,416. If October behaves like the rest of FY27 rather than like FY26's October, the full year gap widens again.

On the five year comparison that was asked for. The pack holds two years of actual revenue: FY27 on the P&L FY27 tab and FY26 on the Actual FY26 tab. FY23, FY24 and FY25 are not in the pack and are not in any file I hold. The columns are in place on the tab, shaded, waiting for a paste, and both charts read whichever year columns carry figures, so nothing has to be rebuilt once the exports arrive.

What is needed is one Xero profit and loss for each of FY23, FY24 and FY25, set to monthly columns, total income by month. Those three figures per year are all the charts use."""

B_BULLETS = [
    'First quarter revenue is $1,624k against $1,229k last year, up $396k or 32%.',
    'Support contributed $719k of that, Integration $441k and Events $420k.',
    'The growth is real on revenue. It did not reach the result, because the '
    'margin delivered on it was 38.9%.',
    'FY23 to FY25 are not in the pack. One Xero profit and loss by month for '
    'each year is all that is needed.',
]
B_NOTES = """First quarter revenue year on year, taken from the Cost Analysis tab.

FY27 first quarter $1,624,460 against FY26 at $1,228,582, which is $395,878 or 32.2% ahead.

By department for the FY27 quarter: Support $718,587, Integration $441,414, Events $420,121, Video $40,270 and Consulting $4,068.

The growth is genuine on revenue and it is worth saying so, because the rest of this pack is about the shortfall against budget. The business is billing a third more than it did a year ago.

What it did not do is reach the result. The margin delivered on that revenue was 38.9% for the quarter, and the first quarter still produced only $5,774 of EBITDA. Revenue growth at this margin does not reach the bottom line.

Only FY26 and FY27 are in the pack. The chart shows the years that carry figures, so FY23, FY24 and FY25 will appear as soon as one Xero profit and loss by month is pasted for each of them."""

p = Presentation(sys.argv[1])
w, h, strip = add(p, 6, 7, 'FINANCIAL PERFORMANCE',
                  'Revenue by month, year on year',
                  '/home/user/Claude/deck/snips/rev_years.png',
                  A_BULLETS, A_NOTES, 'snip rev_years')
print(f'page 8: picture {w:.2f} x {h:.2f}, strip {strip:.2f}')
w, h, strip = add(p, 6, 8, 'FINANCIAL PERFORMANCE',
                  'First quarter revenue, year on year',
                  '/home/user/Claude/deck/snips/rev_q1.png',
                  B_BULLETS, B_NOTES, 'snip rev_q1')
print(f'page 9: picture {w:.2f} x {h:.2f}, strip {strip:.2f}')
p.save(sys.argv[2])
