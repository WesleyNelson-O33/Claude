"""The Nov 2026 derivation tab, the same shape as October's.

October's tab explained a missing pay run. November's explains a state split:
the Employment Hero payroll tax report and the Employment Hero split file
agree on total Australian wages and disagree on which state one employee's
wages belong to. The return follows the report, because the report groups by
residential address and that is the basis it is lodged on.

The report figures are typed because they are the source. Everything derived
from them is a formula, so every line can be clicked on and read.
"""
import re

TITLE, HEAD, LABEL = '3', '9', '4'      # styles already in the workbook
MONEY, MONEYB, PLAIN = '1', '10', '0'

# the Employment Hero payroll tax report, 1 to 30 November 2026,
# grouped by residential address
#            state   wages      super     sal sac   pre-tax    ETP
REPORT = [('NSW', 174788.32, 23281.90, 2725.88, 178.62, 875.00),
          ('QLD', 12698.27, 1523.78, 0.0, 0.0, 0.0),
          ('SA', 5616.01, 673.92, 0.0, 0.0, 0.0),
          ('VIC', 20672.04, 2383.45, 0.0, 0.0, 0.0),
          ('WA', 10453.46, 1254.41, 0.0, 0.0, 0.0)]

EXEMPT = [('NSW', 250.00, 'Per diems')]      # exempt pay categories
DAYS, RATE = 30, 0.0545

NOTES = [
    'The split file and the report agree on total Australian wages, so unlike '
    'October there is no pay run missing from either of them.',
    'They disagree on one employee. Nicol Schildknecht is New South Wales in '
    'the split file and Victoria on the report, 462.05 of wages and 55.45 of '
    'super, 517.50 together. Both their November jobs are Melbourne work.',
    'The report groups by residential address, which is the basis the return '
    'is lodged on, so the return follows the report and the employee is '
    'Victorian for November.',
    'Per diems of 250.00 are exempt and are taken out of the New South Wales '
    'gross. On the split file alone, with the per diems left in and the '
    'employee in New South Wales, the month would have been 6,798.41.',
    'Column F of Main & NSW now carries this basis: salaries and wages '
    '177,335.13, superannuation 23,281.90, termination payments 982.69, and '
    'VIC wages 20,672.04 with super 2,383.45.',
    'The two Xero journals follow this tab, so the 11.50 units that employee '
    'worked in November are allocated in the Victorian pot.',
]


def esc(t):
    return str(t).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


class Sheet:
    def __init__(self):
        self.rows = {}

    def put(self, ref, style, kind, value):
        col, row = re.match(r'([A-Z]+)(\d+)', ref).groups()
        self.rows.setdefault(int(row), []).append((col, style, kind, value))

    def text(self, ref, value, style=PLAIN):
        self.put(ref, style, 'text', value)

    def num(self, ref, value, style=MONEY):
        self.put(ref, style, 'number', value)

    def formula(self, ref, value, style=MONEY):
        self.put(ref, style, 'formula', value)

    def xml(self, widths):
        def colnum(c):
            n = 0
            for ch in c:
                n = n * 26 + ord(ch) - 64
            return n
        out = []
        for r in sorted(self.rows):
            body = []
            for col, style, kind, v in sorted(self.rows[r],
                                              key=lambda x: colnum(x[0])):
                ref = f'{col}{r}'
                if kind == 'formula':
                    body.append(f'<c r="{ref}" s="{style}">'
                                f'<f>{esc(v)}</f></c>')
                elif kind == 'number':
                    body.append(f'<c r="{ref}" s="{style}"><v>{v}</v></c>')
                else:
                    body.append(f'<c r="{ref}" s="{style}" t="inlineStr">'
                                f'<is><t xml:space="preserve">{esc(v)}'
                                f'</t></is></c>')
            out.append(f'<row r="{r}">{"".join(body)}</row>')
        cols = ''.join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>'
                       for i, w in enumerate(widths, start=1))
        last = max(self.rows)
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<worksheet xmlns="http://schemas.openxmlformats.org/'
                'spreadsheetml/2006/main">'
                f'<dimension ref="A1:G{last}"/>'
                '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
                '<sheetFormatPr defaultRowHeight="15"/>'
                f'<cols>{cols}</cols>'
                f'<sheetData>{"".join(out)}</sheetData>'
                '<pageMargins left="0.7" right="0.7" top="0.75" '
                'bottom="0.75" header="0.3" footer="0.3"/></worksheet>')


def build():
    s = Sheet()
    s.text('A1', 'November 2026 payroll tax, how the lodged figure was '
                 'arrived at', TITLE)
    s.text('A2', 'Corporate Technology Services Pty Ltd')
    s.text('A3', 'Lodged 6,768.53. Built from the split file alone it would '
                 'have been 6,798.41. This sets out the difference, and the '
                 'Main & NSW tab has been moved on to this basis.')

    s.text('A5', '1. THE EH PAYROLL TAX REPORT, 1 TO 30 NOVEMBER 2026', LABEL)
    for i, h in enumerate(['State', 'Total wages', 'Total super',
                           'Salary sacrifice', 'Pre-tax deductions',
                           'ETP', 'Grand total']):
        s.text(f'{chr(65 + i)}6', h, HEAD)
    for i, (st, w, su, ss, pt, etp) in enumerate(REPORT):
        r = 7 + i
        s.text(f'A{r}', st)
        for j, v in enumerate((w, su, ss, pt, etp)):
            s.num(f'{chr(66 + j)}{r}', v)
        s.formula(f'G{r}', f'=SUM(B{r}:F{r})')
    r = 7 + len(REPORT)
    s.text(f'A{r}', 'Total', HEAD)
    for c in 'BCDEFG':
        s.formula(f'{c}{r}', f'=SUM({c}7:{c}{r - 1})', MONEYB)
    total_row = r

    s.text(f'A{r + 2}', '2. THE EXEMPT PAY CATEGORIES, TAKEN OUT', LABEL)
    s.text(f'A{r + 3}', 'State', HEAD)
    s.text(f'B{r + 3}', 'Amount', HEAD)
    s.text(f'C{r + 3}', 'Category', HEAD)
    for i, (st, amt, what) in enumerate(EXEMPT):
        s.text(f'A{r + 4 + i}', st)
        s.num(f'B{r + 4 + i}', amt)
        s.text(f'C{r + 4 + i}', what)
    ex_first = r + 4
    ex_last = r + 3 + len(EXEMPT)

    base = ex_last + 2
    s.text(f'A{base}', '3. THE WAGES USED IN THE RETURN', LABEL)
    s.text(f'A{base + 1}', 'State', HEAD)
    s.text(f'B{base + 1}', 'Per the report', HEAD)
    s.text(f'C{base + 1}', 'Exempt, taken out', HEAD)
    s.text(f'D{base + 1}', 'Used in the return', HEAD)
    for i, (st, *_rest) in enumerate(REPORT):
        r2 = base + 2 + i
        s.text(f'A{r2}', st)
        s.formula(f'B{r2}', f'=G{7 + i}')
        ex = next((a for t, a, _w in EXEMPT if t == st), 0)
        if ex:
            s.formula(f'C{r2}', f'=-B{ex_first}')
        else:
            s.num(f'C{r2}', 0)
        s.formula(f'D{r2}', f'=B{r2}+C{r2}')
    used_first, used_last = base + 2, base + 1 + len(REPORT)
    r2 = used_last + 1
    s.text(f'A{r2}', 'Total', HEAD)
    for c in 'BCD':
        s.formula(f'{c}{r2}', f'=SUM({c}{used_first}:{c}{used_last})', MONEYB)
    used_total = r2

    t = used_total + 2
    s.text(f'A{t}', '4. THE TAX', LABEL)
    s.text(f'A{t + 1}', 'New South Wales wages')
    s.formula(f'C{t + 1}', f'=D{used_first}')
    s.text(f'A{t + 2}', 'Interstate wages')
    s.formula(f'C{t + 2}', f'=SUM(D{used_first + 1}:D{used_last})')
    s.text(f'A{t + 3}', 'Total Australian wages')
    s.formula(f'C{t + 3}', f'=C{t + 1}+C{t + 2}')
    s.text(f'A{t + 4}', 'Days in November')
    s.num(f'C{t + 4}', DAYS, PLAIN)
    s.text(f'A{t + 5}', 'Threshold entitlement')
    s.formula(f'C{t + 5}',
              f'=ROUND((C{t + 1}/C{t + 3})*(1200000/365*C{t + 4}),2)')
    s.text(f'A{t + 6}', 'Taxable after the threshold')
    s.formula(f'C{t + 6}', f'=C{t + 3}-C{t + 5}-C{t + 2}')
    s.text(f'A{t + 7}', 'Rate')
    s.num(f'C{t + 7}', RATE, PLAIN)
    s.text(f'A{t + 8}', 'PAYROLL TAX PAYABLE', HEAD)
    s.formula(f'C{t + 8}', f'=ROUND(C{t + 6}*C{t + 7},2)', MONEYB)
    s.text(f'A{t + 9}', 'On the split file basis, before the per diems came '
                        'out and before the state move')
    s.num(f'C{t + 9}', 6798.41)
    s.text(f'A{t + 10}', 'Difference, the per diems and the state split')
    s.formula(f'C{t + 10}', f'=C{t + 8}-C{t + 9}')
    s.text(f'A{t + 11}', 'Main & NSW cell F23, which now carries this basis')
    s.formula(f'C{t + 11}', "='Main & NSW'!F23")
    s.text(f'A{t + 12}', 'Agrees, must be nil')
    s.formula(f'C{t + 12}', f'=ROUND(C{t + 8}-C{t + 11},2)', MONEYB)

    w = t + 14
    s.text(f'A{w}', '5. WHY THIS DIFFERS FROM THE SPLIT FILE', LABEL)
    for i, line in enumerate(NOTES):
        s.text(f'A{w + 1 + i}', line)

    c = w + 2 + len(NOTES)
    s.text(f'A{c}', '6. THE CROSS CHECK TO THE SPLIT FILE', LABEL)
    s.text(f'A{c + 1}', 'Split file gross, plus SG super, plus the employer '
                        'contribution')
    s.num(f'C{c + 1}', 257125.06)
    s.text(f'A{c + 2}', 'The report grand total')
    s.formula(f'C{c + 2}', f'=G{total_row}')
    s.text(f'A{c + 3}', 'Difference, must be nil')
    s.formula(f'C{c + 3}', f'=ROUND(C{c + 1}-C{c + 2},2)', MONEYB)
    s.text(f'A{c + 4}', 'The two agree in total. The difference is which '
                        'state 517.50 of it sits in, not how much there is.')

    return s.xml([62, 17, 17, 17, 17, 17, 17])
