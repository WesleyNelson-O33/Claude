"""Better charts for the same three years.

Three columns for one quarter was a table drawn as a chart. Page 9 now carries
all four quarters across the three years, which still answers the first quarter
comparison and adds the three quarters still to come. A page is added for the
rolling twelve months, which is the only honest trend line for a business with
an October and a January like this one.
"""
import sys
sys.path.insert(0, '/home/user/Claude/deck/edits')
from pptx import Presentation
from addslide import add, relay

SNIP = '/home/user/Claude/deck/snips'

Q_BULLETS = [
    'The first quarter was $1,624k, the strongest of the three years.',
    'It is also the smallest quarter of the year. The second quarter carries '
    'about 32% of revenue.',
    'FY26 lost its year in the third quarter: $1,304k against $1,873k, down 30%.',
    'The second quarter is next and it decides the year. FY25 billed $2,277k '
    'in it, FY26 $1,895k.',
]
Q_NOTES = """Revenue by quarter across three years, taken from the Xero History tab.

First quarter: FY25 $1,380,102, FY26 $1,228,582, FY27 $1,624,460.
Second quarter: FY25 $2,276,986, FY26 $1,894,997.
Third quarter: FY25 $1,873,339, FY26 $1,303,789.
Fourth quarter: FY25 $1,469,930, FY26 $1,685,818.

Three things come out of this that the single quarter comparison could not show.

The first is that the quarter just closed is the smallest quarter of the year in both completed years, and FY27 has made it the largest first quarter of the three. That is a genuinely good result and it is worth saying so.

The second is where FY26 actually lost its year. It was the third quarter, January to March, which fell $569,550 or 30.4% below the same quarter of FY25. The first quarter of FY26 was 11.0% below and the second 16.8% below, but the third is where the damage was done. The fourth quarter was the one quarter FY26 beat FY25, by $215,888 or 14.7%.

The third is what is immediately ahead. The second quarter has carried 32.5% of FY25 and 31.0% of FY26, which makes it the quarter the year is won or lost in. October alone was the largest month of both completed years. The FY27 budget has October at $950,416, and the first quarter has run 18.6% behind budget on revenue. A second quarter that behaves like FY26's rather than FY25's would widen the full year gap by around $380,000 on its own.

FY27 shows one quarter because one quarter has closed. The chart fills in as each quarter closes."""

R_BULLETS = [
    'The twelve months to September 2026 billed $6,509k.',
    'The peak was $7,000k in June 2025. The trough was $5,817k in January 2026, '
    'a fall of 17%.',
    'The trend has turned up since January and has recovered $692k of that, '
    'about 12%.',
    'It is still $491k or 7% below the peak. The recovery is real but it is not '
    'finished.',
]
R_NOTES = """Revenue on a rolling twelve month basis, taken from the Xero History tab. Each point is the total of that month and the eleven before it.

This is the chart that takes the seasonality out. For a business where October can be a seven figure month and January can be a quarter of that, the month by month picture and the quarterly picture both move for reasons that have nothing to do with how the business is actually travelling. The rolling twelve months does not.

The peak was $7,000,357 at June 2025, which is the FY25 full year.

It fell from there for seven straight months to $5,817,190 at January 2026, a fall of $1,183,167 or 16.9%. January 2026 is the single worst month in the three years at $268,631 against $918,289 the year before, and it drags the rolling figure down the moment it enters the window.

Since January the line has turned up. It now stands at $6,509,064 at September 2026, which is $691,874 or 11.9% above the trough.

It is still $491,293 or 7.0% below the peak.

That is the honest position. The business is recovering and the direction is right, but it has not yet got back to where it was fifteen months ago. Quarterly growth of 32% measured off the weakest quarter of the three years does not describe a business at a record. This line does.

The line begins at June 2025 because that is the first month with twelve months of Xero history behind it, and it stops at September 2026 because that is the last closed month.

One note on the chart itself. The scale starts at five million rather than at zero. This series only moves between $5.8 million and $7.0 million, and drawn from zero it is a flat line with a 17% fall invisible in it. The axis labels show where the scale starts."""

p = Presentation(sys.argv[1])

w, h, strip = relay(p.slides[8], f'{SNIP}/xero_qtr.png', Q_BULLETS, Q_NOTES,
                    'snip rev_q1', 'Revenue by quarter, year on year')
print(f'page 9: picture {w:.2f} x {h:.2f}, strip {strip:.2f}')

w, h, strip = add(p, 6, 10, 'FINANCIAL PERFORMANCE', 'Revenue, rolling twelve months',
                  f'{SNIP}/xero_roll.png', R_BULLETS, R_NOTES, 'snip xero_roll')
print(f'page 11: picture {w:.2f} x {h:.2f}, strip {strip:.2f}')

p.save(sys.argv[2])
