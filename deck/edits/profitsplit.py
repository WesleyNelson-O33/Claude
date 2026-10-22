"""The profit page comes apart into a monthly one and a quarterly one.

The deck is one file serving both reviews and the irrelevant slides get hidden,
so a chart may never carry the month and the quarter together. The commentary
and the notes split with it: the monthly page talks about the month only and the
quarterly page about the quarter only, because either one may be the one on
screen with the other hidden.
"""
import sys
sys.path.insert(0, '/home/user/Claude/deck/edits')
from pptx import Presentation
from addslide import add, relay

SNIP = '/home/user/Claude/deck/snips'

MONTH_B = [
    'September lost $41k, against a $53k loss last September and a $21k profit '
    'the September before.',
    'This September is $11k better than the same month last year.',
    'Three Septembers and only one above the line, which is FY25.',
    'September revenue was the highest of the three at $562k, so the loss is a '
    'margin problem rather than a revenue problem.',
]
MONTH_N = """Net profit and net loss for September across three years, taken from the Xero History tab. Net profit is the pack's own figure and includes other income.

September 2026 a loss of $41,464. September 2025 a loss of $52,615. September 2024 a profit of $20,737.

This September is $11,152 better than the same month last year, and $62,201 worse than the September before that.

The thing worth drawing out is that the loss is not a revenue failure. September 2026 billed $561,945, which is the highest of the three Septembers, ahead of $356,323 last year and $440,243 the year before. The month earned more than either comparison and still finished below the line.

What changed is the cost of earning it. Cost of sales was $374,130 against $216,749 in September 2024 on revenue only $121,701 higher. The gross margin on the month was 33.4%, against 50.8% in September 2024.

So the question this chart puts to the board is about pricing and delivery cost on the work being won, not about the volume of work being won.

FY25 was pulled from Xero on 22 October 2026. FY26 is the Actual FY26 tab and FY27 the P&L FY27 tab. The month follows the reporting month on Setup, so in October this becomes October. Only the first quarter of FY25 was pulled, so the FY25 bar is blank from October until the rest is pulled."""

QTR_B = [
    'The quarter lost $6k, against a $85k loss last year and a $35k profit the '
    'year before.',
    'That is $80k better than last year and within $6k of breaking even.',
    'It is still $229k below the budget, which had the quarter at a $223k profit.',
    'Quarter revenue was the highest of the three at $1,624k, so again the gap '
    'is margin.',
]
QTR_N = """Net profit and net loss for the first quarter across three years, taken from the Xero History tab. Net profit is the pack's own figure and includes other income.

Q1 2026 a loss of $5,808. Q1 2025 a loss of $85,404. Q1 2024 a profit of $34,595.

The quarter has improved $79,596 on last year and is within $5,808 of breaking even. On that reading the recovery is real and the direction is right.

Against the plan it reads differently. The budget had the first quarter at a $222,700 profit. Coming in at a $5,808 loss is $228,508 below it.

As with the month, this is not a revenue failure. The quarter billed $1,624,460, which is the highest first quarter of the three, ahead of $1,228,582 last year and $1,380,102 the year before. It earned more than either comparison and still finished below the line.

The cost of earning it is where the quarter went. Cost of sales was $992,160 against $708,478 in Q1 2024, which is $283,682 more on revenue $244,358 higher. Overheads were almost unchanged at $641,508 against $638,615. The overhead base is not the problem and has not been for three years. The margin on the work is.

FY25 was pulled from Xero on 22 October 2026. The quarter is July to September, so it is fixed rather than following Setup."""

p = Presentation(sys.argv[1])

slide = p.slides[13]
w, h, strip = relay(slide, f'{SNIP}/profit_month.png', MONTH_B, MONTH_N,
                    'snip profit_line', 'Net profit and net loss, the month')
for sh in slide.shapes:
    if sh.name == 'snip profit_line':
        sh.name = 'snip profit_month'
print(f'page 14: the month | picture {w:.2f} x {h:.2f}, strip {strip:.2f}')

w, h, strip = add(p, 8, 14, 'FINANCIAL PERFORMANCE',
                  'Net profit and net loss, the quarter',
                  f'{SNIP}/profit_qtr.png', QTR_B, QTR_N, 'snip profit_qtr')
print(f'page 15: the quarter | picture {w:.2f} x {h:.2f}, strip {strip:.2f}')

p.save(sys.argv[2])
