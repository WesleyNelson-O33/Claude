"""Three pages for the cost and profit section.

Spend for the month and spend for the quarter, each against budget and the two
years before, with revenue drawn on the line over the top so the periods where
revenue sat under the spend are visible. Then net profit above and below the
line for both, which is the picture that was asked for.

They go in after the cost and profit summary, which is the same order the income
section uses: the summary page first, then the charts that explain it.
"""
import sys
sys.path.insert(0, '/home/user/Claude/deck/edits')
from pptx import Presentation
from addslide import add

SNIP = '/home/user/Claude/deck/snips'

MONTH_B = [
    'September spend was $604k against a budget of $521k, $84k over.',
    'All of the overspend is cost of sales: $374k against $289k budgeted.',
    'Overheads came in at $230k against $232k budgeted, which is on plan.',
    'Spend is up 48% on last September and 44% on the September before.',
]
MONTH_N = """Spend for September across three years and against budget, taken from the Xero History tab.

Spend here is cost of sales plus overheads. September 2026 $604,363. Budget $520,765. September 2025 $408,939. September 2024 $419,910.

September was $83,599 or 16.1% over budget on spend.

Where that sits matters more than the total. Cost of sales was $374,130 against $289,192 budgeted, which is $84,938 over. Overheads were $230,233 against $231,573 budgeted, which is $1,340 under. So the entire overspend is cost of sales and the overheads landed on plan.

That is the better of the two outcomes. Cost of sales moves with revenue, and September billed $561,945 against a budget of $413,665, so some of the extra cost bought extra revenue. Overheads do not move with revenue, and those held.

Against the two earlier years, spend is up $195,425 or 47.8% on September 2025 and $184,453 or 43.9% on September 2024. Revenue over the same two years is up 57.7% and 27.6%, so spend has grown faster than revenue against FY25 and slower than revenue against FY26.

The red line is revenue. It sat above the spend bar in September 2024 and below it in both years since, which is the same thing as saying those two Septembers lost money.

FY25 comes from Xero, pulled on 22 October 2026. FY26 is the Actual FY26 tab and FY27 the P&L FY27 tab. The month follows the reporting month on Setup, so in October this becomes October."""

QTR_B = [
    'Quarter spend was $1,634k against a budget of $1,773k, $140k under.',
    'Overheads have barely moved in three years: $639k, $647k and $642k.',
    'All the movement is cost of sales, up $325k on last year to $992k.',
    'Revenue sat above spend in Q1 2024 and below it in both years since.',
]
QTR_N = """Spend for the first quarter across three years and against budget, taken from the Xero History tab.

Q1 2026 $1,633,668. Budget $1,773,409. Q1 2025 $1,313,986. Q1 2024 $1,347,094.

The quarter is $139,742 or 7.9% under budget on spend. That is the opposite of the month, where spend ran over.

The split is the point. Overheads have been $638,615, $646,721 and $641,508 across the three first quarters, a range of $8,106 on about $640,000. The overhead base has not moved in three years. Every bit of the movement in spend is cost of sales, which went $708,478, $667,265 and $992,160, up $324,895 or 48.7% on last year.

Cost of sales rising with revenue is what should happen. Revenue rose $395,878 over the same year, so cost of sales rose $324,895 to earn it, which is a contribution of about $71,000 on the extra revenue. That is a thin return on the extra work and it is why the quarter is still a loss.

The red line is revenue. In Q1 2024 it sat above the spend bar and the quarter made money. In Q1 2025 and Q1 2026 it sat below and both lost money. The budget line has revenue well clear of spend, which is how it was built to make $223,000.

FY25 comes from Xero, pulled 22 October 2026. The quarter is July to September, so it is fixed rather than following Setup."""

PROFIT_B = [
    'September lost $41k. Last September lost $53k. The September before made '
    '$21k.',
    'The quarter lost $6k, against a $85k loss last year and a $35k profit the '
    'year before.',
    'The quarter is $80k better than last year and close to breakeven.',
    'It is still $229k below budget, which had the quarter at a $223k profit.',
]
PROFIT_N = """Net profit above and below the line, the reporting month and the first quarter across three years, taken from the Xero History tab. Net profit is the pack's own figure and includes other income.

The month: September 2026 a loss of $41,464, September 2025 a loss of $52,615, September 2024 a profit of $20,737.

The quarter: Q1 2026 a loss of $5,808, Q1 2025 a loss of $85,404, Q1 2024 a profit of $34,595.

Two readings come out of this and they point in opposite directions.

The first is that the business has been below the line for two years. Five of the six bars on this chart are losses or near enough, and the only two profits are both FY25. Whatever else the numbers say, September and the first quarter have not made money since FY25.

The second is that the direction is right. The quarter has improved $79,596 on last year and is within $5,808 of breaking even. The month is $11,152 better than the same month last year. On revenue alone the business is billing a third more than it did a year ago.

What stops that being good news is the plan. The budget had the first quarter at a $222,700 profit. Coming in at a $5,808 loss is $228,508 below it. The business has recovered most of the ground it lost in FY26 and is still a long way from the year it was asked to deliver.

FY25 was pulled from Xero on 22 October 2026. Only the first quarter of FY25 was pulled, so the FY25 bar on the monthly side is blank from October onward until the rest is pulled."""

p = Presentation(sys.argv[1])
for at, title, snip, bullets, notes, name in (
    (11, 'Spend for the month, against budget', f'{SNIP}/spend_month.png',
     MONTH_B, MONTH_N, 'snip spend_month'),
    (12, 'Spend for the quarter, against budget', f'{SNIP}/spend_qtr.png',
     QTR_B, QTR_N, 'snip spend_qtr'),
    (13, 'Net profit above and below the line', f'{SNIP}/profit_line.png',
     PROFIT_B, PROFIT_N, 'snip profit_line'),
):
    w, h, strip = add(p, 8, at, 'FINANCIAL PERFORMANCE', title, snip,
                      bullets, notes, name)
    print(f'page {at + 1}: {title} | picture {w:.2f} x {h:.2f}, strip {strip:.2f}')
p.save(sys.argv[2])
