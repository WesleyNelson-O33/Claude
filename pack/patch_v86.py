"""The five cuts carried over from the Q4 FY2026 Business Advisory pack.

They go on Xero History, under the sections already there, because that sheet
is where the pack keeps anything older than the two years the P&L tabs hold.
build_charts.py draws them and the deck snips them, so they are maintained in
the pack like everything else rather than drawn in PowerPoint.

Where a figure is a full year series it is left exactly as the Q4 pack had it.
Danica has said to leave the FY25 profit difference between its own pages as it
stands, so the waterfall opens at 401,710 and the revenue chart shows 368,638
for the same year. Nothing here reconciles the two.

The revenue by type split is the one cut rebuilt from the ledger, because the
Q4 pack moved 161,492 of equipment sales across to equipment hire and the note
on that page asks for the GL version. Every other type ties to the Q4 pie to
the cent.
"""
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

PACK = 'pack/CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Xero History'

NAVY, WHITE, GREY = 'FF242C39', 'FFFFFFFF', 'FFF4F4F4'
BAND = PatternFill('solid', start_color=NAVY, end_color=NAVY)
TOTF = PatternFill('solid', start_color=GREY, end_color=GREY)
RULE = Side(style='thin', color='FFC9CED6')
HEAD = Font(name='Calibri', sz=11, b=True, color=WHITE)
BOLD = Font(name='Calibri', sz=11, b=True, color=NAVY)
BODY = Font(name='Calibri', sz=11, color=NAVY)
NOTE = Font(name='Calibri', sz=10, i=True, color='FF6B7585')
MONEY = '#,##0;[RED]\\(#,##0\\);\\-'
CENTS = '#,##0.00;[RED]\\(#,##0.00\\);\\-'
PCT = '0.0%;[RED]\\(0.0%\\);\\-'

# ---------------------------------------------------------------- the data
# Sticky 3. Profit per employee, read off the Q4 FY2026 pack, page 3.
PER_HEAD = [
    (2022,   9291.19, 22034.94,  -74030.50,  4934.34),
    (2023,  11263.67, 49461.01, -113330.30, 13985.01),
    (2024,   5804.60, 20147.92,  -34467.02,  7357.24),
    (2025,  18222.00,  2994.00,  -15712.00, 12013.00),
    (2026,  15687.00,  6196.00,  -65663.00,  7856.00),
]
# Sticky 8. Full year revenue and net profit, Q4 FY2026 pack, page 8.
FULL_YEAR = [
    (2016, 3099680.10, 226803.52), (2017, 4011690.16, 220314.11),
    (2018, 4142343.45, 705412.84), (2019, 5034111.11, -13425.37),
    (2020, 5135990.77, 671840.84), (2021, 5846071.81, 846696.71),
    (2022, 5484838.46, 212914.18), (2023, 6066065.86, 483402.89),
    (2024, 5904584.79, 265694.40), (2025, 7000356.97, 368637.78),
    (2026, 6113186.02, 235273.19),
]
# Sticky 7. Fourth quarter revenue and net profit, Q4 FY2026 pack, page 7.
Q4_YEAR = [
    (2016,  814577.54,  -92463.84), (2017, 1020322.41,  36001.43),
    (2018, 1076330.51,   53271.00), (2019, 1465027.92, -255543.55),
    (2020,  943093.40,  190433.19), (2021, 1641335.89,   53821.38),
    (2022, 1368334.99,   73855.70), (2023, 1397464.87,  -56750.30),
    (2024, 1619879.19,  -13340.09), (2025, 1469929.74,  -20375.00),
    (2026, 1685817.62,   65971.04),
]
# Sticky 14. Changes in profitability, Q4 FY2026 pack, page 14, as it stands.
WATERFALL = [
    ('FY25 profit',                401710.0,  'total'),
    ('Gross profit',              -152528.0,  'step'),
    ('Gen and admin',                8058.0,  'step'),
    ('Insurances and advertising',  12454.0,  'step'),
    ('Operating expenses',         -32282.0,  'step'),
    ('Employment',                  -2198.0,  'step'),
    ('FY26 profit',                235213.0,  'total'),
]
# Sticky 18. Revenue by type. Xero income by account, accrual, whole of entity.
# FY26 is the twelve months to 30 June 2026, FY27 the three months to 30 Sep 2026.
FY26_ACC = {
    'Contract Support Staff': 2291640.15, 'Equipment Hires': 1090130.22,
    'Production Labour': 783724.25, 'Equipment Sales': 462667.72,
    'Ad Hoc Service Labour': 421925.50, 'Subscriptions & Licences - Income': 241325.91,
    'Conference Call Charge': 215070.10, 'Video Labour': 187099.00,
    'Tech Consulting/Project Manage': 140707.21, 'Installation Labour': 90090.90,
    'Tech Event Management': 73684.48, 'Freight & Travel - PRD': 49562.33,
    'Meal & Travel - PRD': 41012.52, 'Mobile Phone Plans': 12338.75,
    'Contract Maintenance': 10495.00, 'Meal & Travel - VID': 1255.00,
    'Freight & Travel - INTEGRATION': 237.00, 'Freight & Travel - VID': 125.00,
    'Meal & Travel - ONS': 94.99,
}
FY27_ACC = {
    'Contract Support Staff': 606466.07, 'Equipment Sales': 182311.17,
    'Installation Labour': 180250.85, 'Equipment Hires': 172852.77,
    'Production Labour': 172607.75, 'Ad Hoc Service Labour': 94737.72,
    'Subscriptions & Licences - Income': 84810.19, 'Conference Call Charge': 64004.50,
    'Video Labour': 22346.25, 'Meal & Travel - ONS': 14215.01,
    'Tech Event Management': 9545.00, 'Freight & Travel - PRD': 9245.00,
    'Meal & Travel - PRD': 4955.00, 'Mobile Phone Plans': 3112.50,
    'Freight & Travel - INTEGRATION': 1633.00, 'Tech Consulting/Project Manage': 1067.00,
    'Freight & Travel - VID': 300.00,
}
# type -> the income accounts behind it, and whether it is recurring
TYPES = [
    ('Fixed fee contracts', True, ['Contract Support Staff', 'Contract Maintenance',
                                   'Contract Help Desk', 'Mobile Phone Plans']),
    ('Licence fees',        True, ['Subscriptions & Licences - Income']),
    ('Virtual platforms',   True, ['Conference Call Charge']),
    ('Other labour',       False, ['Production Labour', 'Video Labour',
                                   'Installation Labour',
                                   'Tech Consulting/Project Manage',
                                   'Tech Event Management', 'Meal & Travel - ONS',
                                   'Meal & Travel - PRD', 'Meal & Travel - VID']),
    ('Equipment hire',     False, ['Equipment Hires']),
    ('Equipment sales',    False, ['Equipment Sales']),
    ('Ad hoc contracts',   False, ['Ad Hoc Service Labour']),
    ('Miscellaneous',      False, ['Freight & Travel - PRD',
                                   'Freight & Travel - INTEGRATION',
                                   'Freight & Travel - VID']),
]
# What the Q4 pie had, to prove the mapping. Equipment is the one that moves.
# Sticky 16. Ratios, with the benchmark research alongside.
#   gross margin, overheads over gross profit and EBITDA margin are struck from
#   Xero on a comparable basis: FY25 before its depreciation, interest and tax,
#   FY26 which carries none of those.
#   utilisation and the chargeable split are the Q4 FY2026 pack's own figures.
# Benchmarks: SPI Research Professional Services Maturity Benchmark 2026
# (billable utilisation 66.4% average, 75% target, 79.5% top 20%; EBITDA 9.9%
# average, 15% expected), D-Tools integrator data 2026 (gross margin 38.6% in
# 2025), NSCA guidance (integrator overhead in the high twenties per cent of
# revenue). No published benchmark was found for overheads over gross profit
# or for salary to revenue on comparable definitions.
RATIOS = [
    ('Gross margin',              0.3963, 0.4280,
     '38.6% industry 2025',            'Above',        'D-Tools 2026'),
    ('Overheads over gross profit', 0.7898, 0.9101,
     'None published; 100% is breakeven', '9 pts from breakeven', 'CTS'),
    ('EBITDA margin',             0.0936, 0.0407,
     '9.9% average, 15% expected',     'Below range',  'SPI 2026'),
    ('Billable utilisation',      0.79,   0.75,
     '66.4% average, 75% target, 79.5% top 20%', 'At target', 'SPI 2026'),
    ('Overheads over revenue',    0.34,   0.39,
     'High twenties per cent',         'Above',        'NSCA'),
    ('Salary over revenue',       0.48,   0.54,
     'None published',                 'Not comparable', 'CTS'),
    ('Chargeable labour',         0.37,   0.33,
     'Part of salary over revenue',    'Down',         'CTS'),
    ('Non-chargeable labour',     0.11,   0.21,
     'Part of salary over revenue',    'Up',           'CTS'),
]
Q4_PIE = {'Equipment sales': 301175.33, 'Equipment hire': 1251622.61,
          'Virtual platforms': 215070.10, 'Licence fees': 241325.91,
          'Fixed fee contracts': 2314473.90, 'Ad hoc contracts': 421925.50,
          'Other labour': 1317668.35, 'Miscellaneous': 49924.33}


def tot(acc, names):
    return sum(acc.get(n, 0.0) for n in names)


def check():
    """Every type but the two equipment lines must tie to the Q4 pie."""
    moved = []
    for name, _, accounts in TYPES:
        gl_v, pie = tot(FY26_ACC, accounts), Q4_PIE[name]
        if abs(gl_v - pie) > 0.01:
            moved.append((name, gl_v, pie))
    assert {m[0] for m in moved} == {'Equipment hire', 'Equipment sales'}, moved
    assert abs(sum(FY26_ACC.values()) - 6113186.03) < 0.02
    assert abs(sum(FY27_ACC.values()) - 1624459.78) < 0.02
    shift = tot(FY26_ACC, ['Equipment Sales']) - Q4_PIE['Equipment sales']
    print(f'  mapping ties to the Q4 pie on six of eight types')
    print(f'  equipment: the Q4 pack moved {shift:,.2f} from sales to hire')
    return shift


def band(ws, row, width, text):
    ws.cell(row, 1, text).font = HEAD
    for c in range(1, width + 1):
        ws.cell(row, c).fill = BAND


def header(ws, row, names):
    for c, n in enumerate(names, start=1):
        cell = ws.cell(row, c, n)
        cell.font = BOLD
        cell.border = Border(bottom=RULE)
        if c > 1:
            cell.alignment = Alignment(horizontal='right', wrap_text=True)


def main():
    shift = check()
    wb = openpyxl.load_workbook(PACK)
    ws = wb[SHEET]
    rows = {}

    # ---------------------------------------------- 1. profit per employee
    R = 150
    band(ws, R, 5, 'PROFITABILITY PER EMPLOYEE, FY2022 TO FY2026')
    ws.cell(R + 1, 1, 'Carried over from the Q4 FY2026 Business Advisory pack, '
                      'page 3. Profit per head by reporting group.').font = NOTE
    header(ws, R + 2, ['Year', 'ONS', 'PRD', 'CONS', 'CTS'])
    for i, (yr, ons, prd, cons, cts) in enumerate(PER_HEAD):
        r = R + 3 + i
        ws.cell(r, 1, str(yr)).font = BODY
        for c, v in ((2, ons), (3, prd), (4, cons), (5, cts)):
            cell = ws.cell(r, c, round(v, 2))
            cell.font = BODY
            cell.number_format = MONEY
    rows['per_head'] = (R + 2, R + 3, R + 2 + len(PER_HEAD))

    # ----------------------------------------- 2. full year revenue/profit
    R = 160
    band(ws, R, 4, 'FULL YEAR REVENUE AND NET PROFIT, 2016 TO 2026')
    ws.cell(R + 1, 1, 'Carried over from the Q4 FY2026 Business Advisory pack, '
                      'page 8. FY2026 ties to Actual FY26 in this pack.').font = NOTE
    header(ws, R + 2, ['Year', 'Revenue', 'Net profit', 'Net margin'])
    for i, (yr, rev, pft) in enumerate(FULL_YEAR):
        r = R + 3 + i
        ws.cell(r, 1, str(yr)).font = BODY
        for c, v in ((2, rev), (3, pft)):
            cell = ws.cell(r, c, round(v, 2))
            cell.font = BODY
            cell.number_format = MONEY
        cell = ws.cell(r, 4, f'=IF(B{r}=0,"",C{r}/B{r})')
        cell.font = BODY
        cell.number_format = PCT
    rows['full_year'] = (R + 2, R + 3, R + 2 + len(FULL_YEAR))

    # ------------------------------------- 3. fourth quarter revenue/profit
    R = 176
    band(ws, R, 4, 'FOURTH QUARTER REVENUE AND NET PROFIT, 2016 TO 2026')
    ws.cell(R + 1, 1, 'Carried over from the Q4 FY2026 Business Advisory pack, '
                      'page 7. The same cut as above, on a quarter.').font = NOTE
    header(ws, R + 2, ['Year', 'Revenue', 'Net profit', 'Net margin'])
    for i, (yr, rev, pft) in enumerate(Q4_YEAR):
        r = R + 3 + i
        ws.cell(r, 1, str(yr)).font = BODY
        for c, v in ((2, rev), (3, pft)):
            cell = ws.cell(r, c, round(v, 2))
            cell.font = BODY
            cell.number_format = MONEY
        cell = ws.cell(r, 4, f'=IF(B{r}=0,"",C{r}/B{r})')
        cell.font = BODY
        cell.number_format = PCT
    rows['q4_year'] = (R + 2, R + 3, R + 2 + len(Q4_YEAR))

    # -------------------------------------------------- 4. the waterfall
    # openpyxl has no waterfall type, so it is a stacked bar: an invisible
    # base holds each step up at the running total, and the step sits on it.
    R = 192
    band(ws, R, 5, 'CHANGES IN PROFITABILITY, FY2025 TO FY2026')
    ws.cell(R + 1, 1, 'Carried over from the Q4 FY2026 Business Advisory pack, '
                      'page 14, exactly as that pack has it.').font = NOTE
    header(ws, R + 2, ['Step', 'Base', 'Increase', 'Decrease', 'Total'])
    run = 0.0
    for i, (name, v, kind) in enumerate(WATERFALL):
        r = R + 3 + i
        ws.cell(r, 1, name).font = BODY
        if kind == 'total':
            base, inc, dec, total = None, None, None, v
            run = v
        elif v >= 0:
            base, inc, dec, total = run, v, None, None
            run += v
        else:
            run += v
            base, inc, dec, total = run, None, -v, None
        for c, val in ((2, base), (3, inc), (4, dec), (5, total)):
            if val is None:
                continue
            cell = ws.cell(r, c, round(val, 2))
            cell.font = BODY
            cell.number_format = MONEY
    rows['waterfall'] = (R + 2, R + 3, R + 2 + len(WATERFALL))

    # --------------------------------------------- 5. revenue by type
    R = 204
    band(ws, R, 5, 'REVENUE BY TYPE, AVERAGE MONTH')
    for i, t in enumerate([
            'Xero income by account, accrual, whole of entity. FY26 is the '
            'twelve months to 30 June 2026 divided by twelve. FY27 is the '
            'three months to 30 September 2026 divided by three.',
            f'Built from the ledger rather than from the Q4 pack, which moved '
            f'{abs(shift):,.0f} of equipment sales across to equipment hire. '
            f'Every other type ties to that pack to the cent.']):
        ws.cell(R + 1 + i, 1, t).font = NOTE
    hr = R + 3
    header(ws, hr, ['Revenue type', 'FY26 average month',
                    'FY27 average month', 'Change', 'Change %'])
    body_first = hr + 1
    r = body_first
    for name, _, accounts in TYPES:
        a, b = tot(FY26_ACC, accounts) / 12, tot(FY27_ACC, accounts) / 3
        ws.cell(r, 1, name).font = BODY
        for c, v in ((2, a), (3, b), (4, b - a)):
            cell = ws.cell(r, c, round(v, 2))
            cell.font = BODY
            # whole dollars: a chart data label follows the cell's own format,
            # and cents on a monthly average are noise on a slide
            cell.number_format = MONEY
        cell = ws.cell(r, 5, f'=IF(B{r}=0,"",D{r}/B{r})')
        cell.font = BODY
        cell.number_format = PCT
        r += 1
    body_last = r - 1
    rows['by_type'] = (hr, body_first, body_last)

    r += 1
    for label, want in (('Recurring', True), ('Project and transactional', False),
                        ('Total', None)):
        picks = [n for n, rec, _ in TYPES if want is None or rec is want]
        ws.cell(r, 1, label).font = BOLD
        for c, which in ((2, 0), (3, 1)):
            v = sum((tot(FY26_ACC, a) / 12) if which == 0 else (tot(FY27_ACC, a) / 3)
                    for n, _, a in TYPES if n in picks)
            cell = ws.cell(r, c, round(v, 2))
            cell.font = BOLD
            cell.number_format = MONEY
            cell.fill = TOTF
        cell = ws.cell(r, 4, f'=C{r}-B{r}')
        cell.font = BOLD
        cell.number_format = MONEY
        cell.fill = TOTF
        cell = ws.cell(r, 5, f'=IF(B{r}=0,"",D{r}/B{r})')
        cell.font = BOLD
        cell.number_format = PCT
        cell.fill = TOTF
        ws.cell(r, 1).fill = TOTF
        rows['total_' + label.split()[0].lower()] = r
        r += 1

    # ------------------------------------------------- 6. the ratios
    # The Q4 pack's own four ratios, plus three read off the ledger, each
    # against a published benchmark where one exists. Where none does the row
    # says so rather than borrowing a number from a neighbouring industry.
    R = 225
    band(ws, R, 7, 'RATIOS AGAINST INDUSTRY BENCHMARKS')
    for i, t in enumerate([
            'FY25 carries 245,872 of depreciation, 70,110 of interest and '
            '67,606 of income tax. FY26 carries none of them. Every ratio '
            'below is struck before all four so the two years compare.',
            'Utilisation and the chargeable split are from the utilisation '
            'report, as the Q4 FY2026 pack had them. The rest are from Xero.']):
        ws.cell(R + 1 + i, 1, t).font = NOTE
    header(ws, R + 3, ['Ratio', 'FY25', 'FY26', 'Change', 'Benchmark',
                       'Where CTS sits', 'Source'])
    for i, (name, a, b, bench, where, src) in enumerate(RATIOS):
        r = R + 4 + i
        ws.cell(r, 1, name).font = BODY
        for c, v in ((2, a), (3, b)):
            cell = ws.cell(r, c, v)
            cell.font = BODY
            cell.number_format = PCT
        # percentage points, so the difference of two percentages is scaled
        # up by a hundred before the "pts" format is put on it
        cell = ws.cell(r, 4, f'=(C{r}-B{r})*100')
        cell.font = BODY
        cell.number_format = '+0.0" pts";-0.0" pts";"-"'
        for c, v in ((5, bench), (6, where), (7, src)):
            cell = ws.cell(r, c, v)
            cell.font = BOLD if c == 6 else BODY
            cell.alignment = Alignment(horizontal='left', wrap_text=True)
    rows['ratios'] = (R + 3, R + 4, R + 3 + len(RATIOS))
    for col, w in (('A', 34), ('B', 10), ('C', 10), ('D', 11),
                   ('E', 36), ('F', 17), ('G', 15)):
        ws.column_dimensions[col].width = w

    wb.save(PACK)
    print('  sections written to', SHEET)
    for k, v in rows.items():
        print(f'    {k:<16}{v}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
