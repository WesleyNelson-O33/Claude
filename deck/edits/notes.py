"""Rewrite the speaker notes on every slide in a professional register.

The figures are unchanged. What changes is the voice: no "beat", no
instructions to the room, and each slide reads as an explanation with a
recommendation where one is warranted.
"""
import sys
from pptx import Presentation

N = {}

N[1] = """Monthly Business Review for September 2026. Every figure in this pack is drawn from the Financial Controller Pack, which reconciles to the Xero general ledger to the cent for each department and for the company.

Revenue for the month was $561,945 against a budget of $413,665, exceeding budget by $148,280. EBITDA was a loss of $32,437 against a budgeted loss of $107,100.

The month is explained by gross margin and the fixed overhead base rather than by sales performance."""

N[2] = """The review is set out in five sections.

Financial performance opens, as it carries the result for the month. What we spend follows, because the cost base explains most of the outcome. Customers and cash, people and utilisation, and the budget outlook complete the review."""

N[3] = """Where the business stands at September 2026, by department and for the company.

Support generated $243k of revenue at a 41.7% gross margin and $29k after allocated overheads. Events generated $73k at 77.6% and a loss of $54k. Video generated $7k at 79.6% and a loss of $7k. Integration generated $235k at 10.1% and $10k. Consulting generated $4k at 2.1% and a loss of $20k.

For the company, revenue of $562k against a $414k budget, a gross margin of 33.4% against 45.5% in September last year, overheads of $230k against a $232k budget, and EBITDA of a $32k loss against a budgeted $107k loss.

In summary, revenue exceeded budget while the company remained in loss. Support is the only department covering its allocated overheads. The revenue variance arose almost entirely in Integration at a 10.1% gross margin, which contributed turnover with very little profit."""

N[4] = """Financial performance for September 2026 and the quarter to date."""

N[5] = """The P&L summary is taken from the Summary tab of the Financial Controller Pack and is shown as it appears on screen with the outline closed.

Revenue of $561,945 against a budget of $413,665, exceeding budget by $148,280 and 58% above the $356,323 recorded in September last year.

Gross profit of $187,814 at 33.4%, against a budgeted 30.1% and 45.5% in September last year. The budgeted margin for September was unusually low, so exceeding it reflects the budget rather than an improvement in delivery.

Overheads of $230,233 against a budget of $231,573.

EBITDA a loss of $32,437 against a budgeted loss of $107,100, and a net loss of $41,464 after depreciation and amortisation of $9,982.

Year to date, revenue of $1,624,460 against $1,996,110, behind budget by $371,650, with EBITDA of $5,774 against $222,700.

The figure worth carrying from this slide is break even. At the year to date gross margin of 38.9%, the business requires $549,372 of revenue each month to cover overheads and is averaging $541,487. That gap is structural, and additional sales at the current margin would not close it."""

N[6] = """EBITDA against budget, which is the principal risk in this pack.

The quarter produced $5,774 of EBITDA against a budget of $222,700, a shortfall of $216,927 or 97.4%.

The shortfall matters more than the figure alone suggests. The FY27 budget provides for $397,251 of EBITDA for the year and placed $222,700 of that, 56%, in the first quarter. The year was planned to be earned early, and it was not, so the remaining nine months would have to carry the full amount.

To reach $397,251, the nine months from October would need to produce $391,477, or $43,497 a month. The budget asked those nine months for $174,551, or $19,395 a month. That is 2.2 times the monthly profit the plan assumed, in the months the plan treated as the quieter ones.

For context, the whole of FY26 delivered $235,273 of EBITDA. The next nine months would need $391,477, which is 1.7 times a full year of last year's result. In my view the full year EBITDA budget is no longer recoverable, and I would suggest the board is advised of that now rather than at the half year.

The cause is margin rather than sales. Revenue for the quarter reached 81.4% of budget, which is a reasonable result in itself. Gross profit reached 69.3% of budget. The gap between those two percentages is the issue: the business sold reasonably well and retained very little of it. September illustrates the point, with revenue exceeding budget by $148,280 while the month still recorded an EBITDA loss of $32,437.

The cost base cannot absorb that. Overheads ran at $641,508 for the quarter, $213,836 a month, and came in $48,184 under budget. Spending is not the cause and there is little further to remove that would be material. Wages and premises make up most of it and both are committed for the year, so a reduction in retained margin passes straight to the result.

On break even, at the 38.9% gross margin delivered, the business requires $549,372 of revenue a month to cover overheads and is running at $541,487. The business is trading at break even rather than at a profit, in the quarter the budget expected to be the strongest.

Two factors make the next two quarters harder. December is budgeted at a loss because the business closes for two weeks, so the second quarter cannot recover the gap. And with EBITDA at break even there is no earnings buffer, while $303,605 of cash sits in receivables above the 30 day target.

There is a genuine positive. The first quarter of last year was a loss of $85,404, so this quarter is $91,178 better. The business is improving year on year. It is the budget that was set above what the business can deliver, and that distinction is worth drawing clearly.

My recommendations are to reforecast now on margin rather than revenue, to rebuild Consulting around the capacity that exists, to reset the gross margin assumption to the rate delivered over two years, to hold December as budgeted, and to provide the board with a revised full year EBITDA they can rely on."""

N[7] = """Each department against budget, carrying its full overhead charge, taken from section 7 of the Overhead Split tab.

Support: revenue $242,948 against a $248,017 budget. Gross profit $101,290 at 41.7%. Overhead $72,283, of which $59,182 is allocated admin. Result $29,007.
Events: revenue $73,304 against $65,828. Gross profit $56,881 at 77.6%. Overhead $111,108, of which $71,672 is allocated admin. Result a loss of $54,227.
Video: revenue $7,303 against $17,498. Gross profit $5,813 at 79.6%. Overhead $12,659, of which $9,782 is allocated admin. Result a loss of $6,846.
Integration: revenue $234,858 against $59,272. Gross profit $23,756 at 10.1%. Overhead $13,285, of which $4,148 is allocated admin. Result $10,471.
Consulting: revenue $3,531 against $23,050. Gross profit $74 at 2.1%. Overhead $19,944, of which $13,979 is allocated admin. Result a loss of $19,870.

The five departments total a loss of $41,464, which is net profit for the month.

Two observations. Integration sold $175,586 more than planned and it produced very little profit, because the margin retained on resold equipment is 10.1%. And Events carries $71,672 of office costs on $73,304 of sales. That charge is allocated on office usage, and Events is not the department occupying the office. I would suggest the allocation basis is reviewed for FY28 rather than changed mid-year."""

N[8] = """The same month read as three departments rather than five, which is how the Summary tab is constructed.

Production is the parent of Events and Video. Consulting is the parent of Integration and Consulting. Support has no sub-departments. Admin is not a trading department and its costs are allocated across the three.

Support: revenue $242,948 against a $248,017 budget. Gross profit $101,290 at 41.7%. All overhead $72,283. Result $29,007.
Production: revenue $80,607 against $83,326. Gross profit $62,694 at 77.8%. All overhead $123,767. Result a loss of $61,073.
Consulting: revenue $238,389 against $82,322. Gross profit $23,830 at 10.0%. All overhead $33,229. Result a loss of $9,398.

The three total a loss of $41,464, which is net profit for the month, and the check row on the tab reconciles it to the five department view.

The grouping shows something the five departments did not. Production, being Events and Video together, retains 77.8% of what it sells, the strongest margin in the business. It records a loss only because it carries $123,767 of running costs against $80,607 of sales. That is a volume position rather than a pricing one, and it presents the clearest growth case in the pack.

The Consulting group is the opposite. It sold $156,068 more than plan and retains 10.0%, because almost all of that was resold equipment through Integration. Additional sales of that kind would not improve the result.

Support carries the profit of the business. On these three lines the questions are how Support is protected, how Production is given volume, and how Consulting is resolved."""

N[9] = """Support for September 2026, taken from the Dept P&L tab.

Revenue $242,948 against a budget of $248,017, behind by $5,068. September last year was $231,214.

After direct costs of $141,659 it retained $101,290, being 41.7% of revenue.

Overheads charged to it were $72,283, of which $13,101 is its own spending and $59,182 is its share of the office, allocated on the keys set in the pack.

That leaves $29,007 for the month.

Year to date it has billed $718,587 against a budget of $743,185.

On people, employment expense of $51,242, of which $8,980 is its own staff and $42,262 its share of office wages. It recorded 2,132 billable hours and 11 non-billable, with 218 hours of leave."""

N[10] = """Events for September 2026, taken from the Dept P&L tab.

Revenue $73,304 against a budget of $65,828, ahead by $7,476. September last year was $75,333.

After direct costs of $16,423 it retained $56,881, being 77.6% of revenue.

Overheads charged to it were $111,108, of which $39,436 is its own spending and $71,672 is its share of the office.

That leaves a loss of $54,227 for the month.

Year to date it has billed $420,121 against a budget of $321,957.

On people, employment expense of $50,959, of which $35,665 is its own staff and $15,294 its share of office wages. It recorded 241 billable hours and 249 non-billable, with 333 hours of leave.

The department earns the second strongest margin in the business and still records a loss, because the overhead allocated to it is larger than its revenue. I would suggest the allocation basis is the item to examine here, alongside the leave taken in the month."""

N[11] = """Video for September 2026, taken from the Dept P&L tab.

Revenue $7,303 against a budget of $17,498, behind by $10,196. September last year was $20,025.

After direct costs of $1,489 it retained $5,813, being 79.6% of revenue.

Overheads charged to it were $12,659, of which $2,878 is its own spending and $9,782 is its share of the office.

That leaves a loss of $6,846 for the month.

Year to date it has billed $40,270 against a budget of $85,584.

On people, employment expense of $3,961, of which $2,660 is its own staff and $1,301 its share of office wages. It recorded 25 billable hours and 8 non-billable, with no leave.

The margin is the highest in the business. The constraint is the volume of billable work, which at 25 hours for the month is well below the level the budget assumed."""

N[12] = """Integration for September 2026, taken from the Dept P&L tab.

Revenue $234,858 against a budget of $59,272, ahead by $175,586. September last year was $21,421.

After direct costs of $211,102 it retained $23,756, being 10.1% of revenue.

Overheads charged to it were $13,285, of which $9,136 is its own spending and $4,148 is its share of the office.

That leaves $10,471 for the month.

Year to date it has billed $441,414 against a budget of $608,676.

On people, employment expense of $8,832, of which $8,563 is its own staff and $269 its share of office wages. It recorded 10 billable hours and no non-billable time.

The revenue came from a single completed installation rather than from recurring work, and from equipment supply rather than labour, which is why 10 billable hours sit behind $235k of sales. I would treat the result as timing rather than as a change in run rate."""

N[13] = """Consulting for September 2026, taken from the Dept P&L tab.

Revenue $3,531 against a budget of $23,050, behind by $19,519. September last year was $8,330.

After direct costs of $3,457 it retained $74, being 2.1% of revenue.

Overheads charged to it were $19,944, of which $5,965 is its own spending and $13,979 is its share of the office.

That leaves a loss of $19,870 for the month.

Year to date it has billed $4,068 against a budget of $236,707.

On people, employment expense of $9,496, of which $5,545 is its own staff and $3,951 its share of office wages. It recorded 31 billable hours and 84 non-billable, with 53 hours of leave.

The department has carried running costs against almost no fee income for three consecutive months. I would suggest its structure, capacity and commercial basis are reviewed, including whether the work is better delivered within Integration."""

N[14] = """Revenue by department for September 2026, actual against budget and against the same month last year.

Support: $242,948 against a budget of $248,017 and $231,214 last year.
Integration: $234,858 against a budget of $59,272 and $21,421 last year.
Events: $73,304 against a budget of $65,828 and $75,333 last year.
Video: $7,303 against a budget of $17,498 and $20,025 last year.
Consulting: $3,531 against a budget of $23,050 and $8,330 last year.

Company revenue of $561,945 against a $413,665 budget and $356,323 in September last year.

The variance to budget is almost entirely Integration, and it represents one large fit-out at a 10.1% gross margin. It is better read as timing than as growth. Video and Consulting are below budget for the third consecutive month."""

N[15] = """Running costs are steady at approximately $214k a month. Spending is not the cause of the result, but it is the figure revenue has to cover."""

N[16] = """Operating cost distribution for the year to date, with cost of sales and overhead shown on one scale.

Cost of sales, labour $620,737 at 38.0%. Cost of sales, equipment and services $347,307 at 21.3%. Cost of sales, other direct $24,117 at 1.5%. Overhead, general and administrative $75,700 at 4.6%. Overhead, business insurances $25,227 at 1.5%. Overhead, advertising and promotion $34,119 at 2.1%. Overhead, operating expenses $136,752 at 8.4%. Overhead, employment expenses $369,466 at 22.6%.

The concentration is the point. Wages for staff on jobs and wages for everyone else make up the bulk of the cost base, and the second group is largely fixed within the year, which is why a quieter month passes straight to the result.

If the question is where costs can be reduced, the honest answer is headcount or premises. Reductions in subscriptions and travel would not change this picture materially."""

N[17] = """Other operating expenses for the year to date.

Rent $62,043 at 45.4%. IT network service and support $39,570 at 28.9%. Rental outgoings $14,684 at 10.7%. Travel and per diems $5,627 at 4.1%. Stamp duty, indirect, $3,455 at 2.5%. Repairs and maintenance $2,079 at 1.5%. Office cleaning $2,079 at 1.5%. Mobile phones $1,825 at 1.3%. All other items $5,390 at 3.9%.

Rent and outgoings are fixed by the lease. IT support is the next largest and the only item that moves with activity.

The practical conclusion is that savings of any size do not come from this category. A material reduction would require a discussion about the office lease."""

N[18] = """Labour mix for the year to date, direct labour charged to cost of sales split by type.

Sub-contract labour $117,012. Service hires $23,704. Direct salaries $314,763. Direct wages $111,018. Direct superannuation $54,240.

The majority is permanent payroll. Contractors are the only labour cost that rises and falls with the work.

This is the link to the utilisation section. The business pays for hours worked and invoices for hours that are chargeable, and margin follows the difference between the two."""

N[19] = """Employment expense split between office staff and departmental staff for September 2026.

Departmental staff is the employment expense on each department's own cost centre: Support $8,980, Events $35,665, Video $2,660, Integration $8,563, Consulting $5,545.

Office staff is the admin employment pot of $63,077, allocated across the departments on their share of full time equivalent: Support $42,262, Events $15,294, Video $1,301, Integration $269, Consulting $3,951.

Full time equivalent is rolling twelve month worked hours on that cost centre divided by 2,024 hours, being 253 working days at eight hours. That is the method set out in the training material.

The change worth noting is that this allocation no longer depends on a figure typed in by hand. The pack calculates it from the Labour Hours Report, so it reproduces each month without a judgement being made."""

N[20] = """Revenue is concentrated and cash is slow. Days sales outstanding is 47.2 days against a 30 day target, which represents $303,605 of cash held in the receivables book above where it should be."""

N[21] = """The ten largest customers by revenue for the year to date, with the top three highlighted.

1. Deloitte Services: $401,755 for the year, 24.7% of the total, $119,614 of it in September.
2. Aware Super: $285,075 for the year, 17.5% of the total, $241,178 of it in September.
3. CBA Place South: $105,040 for the year, 6.5% of the total, $30,409 of it in September.
4. Ventia: $97,662 for the year, 6.0% of the total, $9,105 of it in September.
5. PwC Services: $88,051 for the year, 5.4% of the total, $26,707 of it in September.
6. APT Management Services: $63,613 for the year, 3.9% of the total, $4,527 of it in September.
7. ASX Operations: $60,858 for the year, 3.7% of the total, $17,167 of it in September.
8. Bankwest: $54,971 for the year, 3.4% of the total, $17,230 of it in September.
9. Data#3: $45,909 for the year, 2.8% of the total, none in September.
10. Generation-e Productivity Solutions: $37,166 for the year, 2.3% of the total, $12,389 of it in September.

The ten together represent $1,240,100, or 76.3% of total revenue. Every other customer accounts for the remaining $384,360. The three highlighted customers are 48.7% of the year between them.

The distinction that matters is contracted work against one-off projects. Deloitte Services is contracted support that renews monthly. Aware Super is a fit-out that has now completed. Treating the two as the same kind of income is how a forecast goes wrong.

I would suggest a named relationship owner for each of the top three before the half year, and that the new business target is set on customers four to twenty rather than on the accounts the business already holds."""

N[22] = """Days sales outstanding at 30 September 2026.

Receivables are $833,320. Over the last three months the business sold $1,624,460, which across 92 days is $17,657 of sales a day. On that basis customers are taking 47.2 days to pay.

The target is 30 days, so the book is 17.2 days over. Collected on time, receivables would be $529,715 rather than $833,320, leaving $303,605 of cash in the book that should already be with the business.

Set against the result, the quarter produced $5,774 of EBITDA while $303,605 remains uncollected. Collection is worth considerably more this quarter than any trading improvement available in the time, and it requires no additional cost.

Each day removed from the figure returns approximately $17,657."""

N[23] = """The ten largest amounts owed at 30 September 2026.

Aware Super: $277,395, 33.3% of the total.
Deloitte Services P/L for Deloitte Services Trust: $131,575, 15.8%.
Ventia: $94,701, 11.4%.
Downer EDI Limited: $39,664, 4.8%.
Generation-e Productivity Solutions: $39,106, 4.7%.

The ageing profile is current $588,782 at 70.7%, under one month $133,177 at 16.0%, one month $86,553 at 10.4%, and three months $24,809 at 3.0%. Overdue in total $244,538, or 29.3% of the book.

Two customers hold most of the balance, so a targeted approach to those accounts would be more productive than a general collections exercise, and it represents the fastest cash available to the business this quarter."""

N[24] = """The company invoiced 2,438 hours in September 2026 out of 3,395 it was available for, which is 72%. Where chargeable hours fall short, margin follows."""

N[25] = """Chargeable per cent for September 2026, calculated on the same basis as the monthly Utilisation Report.

The measure is billable hours divided by workable business hours, being the hours each person was available for in the month. Leave and public holidays remain in the denominator, so a team taking its leave entitlement shows a lower percentage. That is the business's own definition and the pack now follows it.

Support 91.2% against an 88.0% target, 3.2 points over. Production 30.2% against 67.0%, 36.8 points under. Consulting 23.0% against 50.0%, 27.0 points under. The company 71.8% against 80.9%, 9.0 points under.

Support has exceeded its target in all three closed months, by 1.4, 5.8 and 3.2 points.

Leave explains the Production position: 333 hours out of 882 available, being 37.7% of the month. Across the company, leave accounted for 17.8% of available hours against 7.5% in August.

The report shows three person groups rather than five, because it groups by the department a person belongs to and nobody is assigned to Video or Integration. Their work is carried out by Production and Consulting staff.

On sources, workable business hours and the target are both entered from the report, as neither can be derived in the pack: the first requires each group's headcount and part time arrangements, and the second is set per person and averaged. The over or under reading and the leave share are calculated and reproduce the report's own columns. The three month history is held on the Utilisation tab with two reconciliation rows against the report's grand total.

One correction to note for completeness. An earlier version of this pack divided billable hours by hours actually worked, excluding leave, which showed Support at 99.5% and gave the appearance of a coding error. That was a difference in method rather than a problem with the data.

OT and TIL is reported but is not carried in the pack. The August figure for Support reads 365.62, which cannot represent hours in a month, so that column would benefit from review on the source report before it is relied upon."""

N[26] = """How each department spent its hours in September 2026.

The bars show the mix within each department rather than its size. Each is scaled to that department's own total, because Support works around a hundred times the hours Integration does and on a common scale four of the five would not be visible. The figure on each segment is the actual hours.

Support: 2,132 hours on jobs, 11 not on jobs, 218 leave.
Events: 241 hours on jobs, 249 not on jobs, 333 leave.
Video: 25 hours on jobs, 8 not on jobs, no leave.
Integration: 10 hours on jobs, none off jobs, no leave.
Consulting: 31 hours on jobs, 84 not on jobs, 53 leave.

Across the five departments, 2,437 hours on jobs out of 2,789 worked.

On the Support bar, the small non-billable segment arises from the regrouping rather than from coding. The monthly report shows Support with no non-billable hours in September. This view groups by the department the work was performed for, so a small amount of other staff's non-billable time appears here.

Events is the clearest operational position on this slide: 333 hours of leave on a small team removes most of the month's capacity.

There is also an opportunity worth quantifying. Support earns approximately $48 of margin on each hour it invoices. Five additional points of billable time across Support and Events, at the rates each already achieves, would be worth approximately $15,813 a month at no additional cost. It is the largest controllable item in this pack."""

N[27] = """The hours mix for September 2026, grouped as the three departments the P&L is built on.

Production is Events and Video together. Consulting is Integration and Consulting together. Support stands alone.

Each bar is scaled to its own total, as Support works considerably more hours than the other two groups combined.

The pattern underlies much of this pack. Support is almost entirely billable and is the only group covering its costs. Production and Consulting each spend around half their paid hours, or more, on work that cannot be invoiced, and both record losses.

On the monthly report's own basis, which counts leave against the measure, Support reads 91.2%. That is the figure to use in discussion."""

N[28] = """Three months are closed and the remaining nine remain at original budget. That is the matter before the board."""

N[29] = """Revenue by month, actual for the three closed months and original budget thereafter.

First quarter actual of $1,624,460 against a budget of $1,996,110, behind by $371,650.

The nine months from October remain at the budget set before the year began. October is budgeted at $950,416, which is higher than any month the business has delivered.

Two points for the board. The first quarter variance is concentrated in Consulting and Video, which were budgeted on a resourcing level that does not exist. And the margin assumption in the budget sits above the rate the business has delivered for two years, so a month that meets the revenue figure would still fall short on profit.

My recommendation is a rolling forecast from the second quarter, rebuilding Consulting around actual capacity and resetting gross margin to delivered rates. December would be better left as budgeted, as the business closes for two weeks and the loss is genuine."""

N[30] = """Seven recommendations, each with a suggested owner and date.

The first two carry the most weight: reforecasting on margin rather than revenue, and reviewing the overhead keys. Most of the pack follows from those two figures.

The Consulting position has now run for three months with running costs against almost no fee income, and would benefit from a decision this quarter.

The receivables item is the only one that produces cash within the quarter, and utilisation remains the largest controllable item."""

N[31] = """The source files behind this pack. Every figure in the review is drawn from the Financial Controller Pack, which reconciles to the Xero general ledger."""

N[32] = """Two figures to close on. At the margin the business delivers, break even requires $549,372 of revenue a month, and the business is running at $541,487. Everything else in the pack sits underneath those two numbers."""

p = Presentation(sys.argv[1])
n = 0
for i, s in enumerate(p.slides, 1):
    if i not in N:
        continue
    tf = s.notes_slide.notes_text_frame
    tf.text = N[i]
    n += 1
p.save(sys.argv[2])
print(f'notes rewritten on {n} slides')
