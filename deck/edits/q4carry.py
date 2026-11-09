"""The six pages carried over from the Q4 FY2026 Business Advisory pack.

Each one is a snip of a chart or table now living in the controller pack, on
Xero History, so they are maintained with the rest of the pack rather than
drawn in PowerPoint.

Where they go:
  11  revenue by type, average month          sticky 18
  17  revenue and net profit, full year       sticky 8
  18  revenue and net profit, fourth quarter  sticky 7
  19  changes in profitability                sticky 14
  45  profit per employee                     sticky 3
  46  ratios against industry benchmarks      sticky 16

The full year page and the quarter page are the same cut on two bases, so they
sit next to each other and either can be hidden without the other reading as a
dangling reference.
"""
import sys
from pptx import Presentation
from addslide import add

SNIP = '/home/user/Claude/deck/snips/'
FIN, PEOPLE = 'FINANCIAL PERFORMANCE', 'PEOPLE AND UTILISATION'

PAGES = [
 # (index to insert at, eyebrow, title, snip, bullets, notes)
 (40, PEOPLE, 'Ratios against industry benchmarks', 'q4_ratios', [
    'Gross margin is 42.8%, above the 38.6% the integrator benchmark reports '
    'for 2025.',
    'EBITDA margin is 4.1% against a 9.9% professional services average, and '
    'against the 15% that benchmark treats as the expected level.',
    'Overheads take 91% of gross profit. At 100% the company breaks even, so '
    'there are nine points of cover left.',
    'Utilisation is 75%, which is the benchmark target and nine points above '
    'the 66.4% average.',
    'Salary is 54% of revenue, up six points, and non-chargeable labour has '
    'nearly doubled from 11% to 21%.',
    'No published benchmark was found for overheads over gross profit or for '
    'salary to revenue on comparable definitions.',
  ], """Ratios against published benchmarks, from the Ratios page of the Q4 FY2026 Business Advisory pack, with the benchmark research added.

    The comparability point first, because it changes how the year on year column reads. FY25 carries $245,872 of depreciation, $70,110 of interest and $67,606 of income tax. FY26 carries none of them: no depreciation or amortisation has been booked at all this year, and the only account in that family is Instant Asset Write Off. Every ratio on this page is struck before all four items so the two years sit on the same basis. The net profit figures on the revenue and profit pages are not adjusted that way, so do not read the two sets against each other.

    Gross margin 39.6% to 42.8%. D-Tools integrator data puts industry gross margin at 38.6% in 2025, down from 41.9% in 2024. CTS is above that and improving, which is the one clearly good line on the page.

    Overheads over gross profit 79.0% to 91.0%. That is twelve points worse in one year and it is the ratio that decides whether the company makes money. No published benchmark was found on a comparable definition, so the reference line is breakeven at 100%.

    EBITDA margin 9.4% to 4.1%. The SPI Research Professional Services Maturity Benchmark 2026 puts the average at 9.9% and the expected level at 15%. CTS was at the average last year and is now well under half of it.

    Billable utilisation 79% to 75%. The same benchmark reports 66.4% average, 75% target and 79.5% for the top fifth. CTS is on target and above average, so utilisation is not where the margin is going.

    Salary to revenue 48% to 54%, with chargeable labour down from 37% to 33% and non-chargeable up from 11% to 21%. Utilisation holding at target while non-chargeable labour nearly doubles means the extra cost is in roles that were never chargeable, not in chargeable staff sitting idle.

    Sources: SPI Research Professional Services Maturity Benchmark 2026; D-Tools integrator benchmark data 2026; NSCA guidance that integrator overhead typically sits in the high twenties per cent of revenue. Utilisation and the chargeable split are the Q4 pack's own figures from the utilisation report."""),

 (40, PEOPLE, 'Profit per employee', 'q4_perhead', [
    'ONS earned $15,687 a head, down from $18,222 the year before.',
    'PRD recovered to $6,196 from $2,994, but is still a long way under the '
    '$49,461 it reached in 2023.',
    'CONS lost $65,663 a head, its worst year since 2023 and four times the '
    'loss it made in 2025.',
    'CTS overall earned $7,856 a head, down from $12,013.',
    'CONS has not earned a positive figure in any of the five years on this '
    'chart.',
  ], """Profit per employee by reporting group, FY2022 to FY2026, carried over from page 3 of the Q4 FY2026 Business Advisory pack.

    ONS 9,291 then 11,264 then 5,805 then 18,222 then 15,687.
    PRD 22,035 then 49,461 then 20,148 then 2,994 then 6,196.
    CONS -74,031 then -113,330 then -34,467 then -15,712 then -65,663.
    CTS 4,934 then 13,985 then 7,357 then 12,013 then 7,856.

    The axis has to carry CONS at -113,330 and PRD at 49,461 on the same scale, which bunches the top three lines together. That bunching is the finding rather than a fault in the chart: ONS, PRD and CTS have all sat inside a band of roughly 3,000 to 22,000 a head for four of the five years, while CONS swings by tens of thousands.

    CONS has never been positive on this measure. FY2025 was the closest it came, at -15,712, and FY2026 gave back four fifths of that recovery.

    FY2027 is not on the chart. Three months of a year cannot sit on a chart of full years without reading as a collapse that has not happened. Year to date the three groups stand at ONS 70,736, PRD -53,072 and CONS -26,871 in total profit, not per head."""),

 (15, FIN, 'Changes in profitability, FY2025 to FY2026', 'q4_bridge', [
    'Profit fell from $401,710 to $235,213, down $166,497.',
    'Gross profit is the whole of it, down $152,528.',
    'Operating expenses took another $32,282 and employment $2,198.',
    'Gen and admin gave back $8,058 and insurances and advertising $12,454.',
    'The two gains together recover $20,512, which is an eighth of the fall.',
    'Nothing on the overhead side of this bridge comes close to the gross '
    'profit movement.',
  ], """Changes in profitability FY2025 to FY2026, carried over from page 14 of the Q4 FY2026 Business Advisory pack, exactly as that pack has it.

    FY25 profit 401,710.
    Gross profit -152,528.
    Gen and admin +8,058.
    Insurances and advertising +12,454.
    Operating expenses -32,282.
    Employment -2,198.
    FY26 profit 235,213.

    One thing to know before this is presented. The Q4 pack does not agree with itself on the FY25 opening figure. This bridge starts at 401,710; the revenue and profit page in the same pack shows 368,638 for FY25. That is 33,072 apart. Danica has asked for both to stay as they are, so neither has been changed here.

    The shape of the bridge is the point regardless of which opening figure is right. Of the 166,497 fall, gross profit accounts for 152,528, which is 92% of it. The four overhead movements net to 13,968 against the company, so the overhead side of the business was close to flat.

    That matters for where effort goes. A gross profit problem of this size is a pricing and mix problem, not a cost control problem, and the overhead lines on this page confirm that there was no cost blowout to point at."""),

 (15, FIN, 'Revenue and net profit, fourth quarter', 'q4_quarter', [
    'Fourth quarter revenue was $1,685,818, the highest fourth quarter on this '
    'chart.',
    'Fourth quarter profit was $65,971, the first positive one since 2022.',
    'Four of the five fourth quarters before this one lost money or made less '
    'than 1%.',
    'The worst was 2019, which lost $255,544 on $1,465,028 of revenue.',
    'The best was 2020 at 20%, on the lowest revenue of the eleven years.',
  ], """Fourth quarter revenue and net profit, 2016 to 2026, carried over from page 7 of the Q4 FY2026 Business Advisory pack.

    This is the quarterly half of a pair. The page before it is the same cut on a full year. Hide either one and the other still reads on its own, because neither refers across to the other.

    Revenue by fourth quarter, 2016 to 2026: 814,578; 1,020,322; 1,076,331; 1,465,028; 943,093; 1,641,336; 1,368,335; 1,397,465; 1,619,879; 1,469,930; 1,685,818.

    Profit by fourth quarter: -92,464; 36,001; 53,271; -255,544; 190,433; 53,821; 73,856; -56,750; -13,340; -20,375; 65,971.

    The fourth quarter has been the weak quarter for most of this history. Three of the eleven lost money and two more made under 1%. FY2026 breaking that run at 4% is the better news on this page, and it comes on the highest fourth quarter revenue in the series.

    Revenue and the quarter's profit do not move together. 2020 made the highest margin of the eleven years on the lowest revenue, and 2019 made the worst loss on what was then the highest revenue."""),

 (15, FIN, 'Revenue and net profit, full year', 'q4_fullyear', [
    'FY2026 revenue was $6,113,186, down $887,171 on FY2025.',
    'FY2026 net profit was $235,273, against $368,638 the year before.',
    'Net margin has sat between 4% and 5% in four of the last five years.',
    'The two strong years on this chart are 2018 at 17% and 2021 at 14%.',
    'Revenue has roughly doubled since 2016 while profit has not.',
    'FY2026 ties to Actual FY26 in the controller pack to the cent.',
  ], """Full year revenue and net profit, 2016 to 2026, carried over from page 8 of the Q4 FY2026 Business Advisory pack.

    This is the full year half of a pair. The page after it is the same cut on a fourth quarter.

    Revenue: 3,099,680; 4,011,690; 4,142,343; 5,034,111; 5,135,991; 5,846,072; 5,484,838; 6,066,066; 5,904,585; 7,000,357; 6,113,186.

    Net profit: 226,804; 220,314; 705,413; -13,425; 671,841; 846,697; 212,914; 483,403; 265,694; 368,638; 235,273.

    FY2026 revenue of 6,113,186 and profit of 235,273 both agree with Actual FY26 in the controller pack to the cent, so this page and the pack's own P&L are the same numbers.

    The shape to notice is that revenue has nearly doubled over eleven years while profit has not moved with it. The margin line sits in a four to five per cent band in seven of the eleven years, with 2018 and 2021 the exceptions.

    One caution on comparing FY25 and FY26 on this page. FY25 carries 245,872 of depreciation, 70,110 of interest and 67,606 of income tax. FY26 carries none of them. The two bars are not struck on the same basis. The ratios page later in the deck restates both years before all four items, and on that basis EBITDA margin falls from 9.4% to 4.1%.

    The Instant Asset Write Off note on the original version of this page said that removing the write off would lift net margin to 9.3% in 2023, 5.3% in 2024, 6.1% in 2025 and 5.0% in 2026, on asset spend of 30k, 47k, 25k and 71k."""),

 (10, FIN, 'Revenue by type, average month', 'q4_revtype', [
    'Revenue averaged $541,487 a month against $509,432 last year.',
    'Recurring income grew $21,892 a month, project work $10,163.',
    'Equipment hire fell $33,227, the largest single move on the page.',
    'Equipment sales rose $22,215 and other labour $25,190.',
  ], """Revenue by type, FY26 average month against FY27 year to date average month, carried over from page 18 of the Q4 FY2026 Business Advisory pack.

    Built from the ledger rather than from that pack, because the note on the original page asks for the GL version. The Q4 pack moved 161,492 of equipment sales across to equipment hire. Every other type ties to that pack to the cent.

    FY26 is the twelve months to 30 June 2026 divided by twelve. FY27 is the three months to 30 September 2026 divided by three. Three months is a short base and August is results season, so treat the monthly averages as indicative.

    Fixed fee contracts 192,873 to 203,193, up 10,320.
    Licence fees 20,110 to 28,270, up 8,160.
    Virtual platforms 17,923 to 21,335, up 3,412.
    Other labour 109,806 to 134,996, up 25,190.
    Equipment hire 90,844 to 57,618, down 33,227.
    Equipment sales 38,556 to 60,770, up 22,215.
    Ad hoc contracts 35,160 to 31,579, down 3,581.
    Miscellaneous 4,160 to 3,726, down 434.

    Recurring 230,906 to 252,798, up 21,892 or 9.5%.
    Project and transactional 278,526 to 288,689, up 10,163 or 3.6%.
    Total 509,432 to 541,487, up 32,054 or 6.3%.

    The recurring base is now 47% of revenue and growing faster than the project side, which is the direction the business has said it wants. The equipment lines are the volatile ones: hire down a third and sales up by over half, on a three month base."""),
]


def main(src, out):
    pres = Presentation(src)
    before = len(pres.slides)
    for at, eyebrow, title, snip, bullets, notes in PAGES:
        w, h, strip = add(pres, 5, at, eyebrow, title, SNIP + snip + '.png',
                          bullets, notes, 'snip ' + snip)
        print(f'  {snip:<12} at index {at:<3} picture {w:.2f} x {h:.2f}, '
              f'strip {strip:.2f}')
    pres.save(out)
    print(f'{before} slides in, {len(pres.slides)} out')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
