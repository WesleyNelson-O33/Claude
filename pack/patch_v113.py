"""Put the rolling forecast utilisation schedule into the controller pack.

Danica sent the pivot and asked for it in the pack, with every column her
pivot carries, not the shortened set the Q1 slide uses. So this one keeps
Sum of Public Holiday, Total Hours incl of PH, Non-chargeable % and % of
Leave alongside the columns already on 'Utilisation Q1'.

The hours, the public holiday and each person's target are typed, because
they are what she sent. Everything else is a formula off them, so a changed
hour moves the percentages on its own:

    Total Hours incl of PH   workable + public holiday
    Chargeable %             chargeable / workable
    Non-chargeable %         1 - chargeable %
    Over / Under             chargeable % - target
    % of Leave               leave / workable

Checked against her pivot before this was written. Every percentage comes
back the same. Shown to one decimal, as 'Utilisation Q1' does, because her
pivot rounds to whole percent and three cells read a point out against it. The group totals are SUM of the rows above, which differ from
her pivot subtotals by a hour or two because the per person hours she sent
are rounded to whole hours. No percentage moves.

The conditional formatting is her own: a green data bar across Over / Under,
colour 63C384, the same rule patch_v108 put on the Q1 sheet.

Run it with no arguments.
"""
# Every Font here names Calibri. Without a name openpyxl leaves the workbook
# default, which comes out of LibreOffice as a serif and does not match the
# rest of the deck.
import openpyxl
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

PACK = 'CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Utilisation Rolling'
TOP = 1
INK, TEAL, BAND, LINE = '242C39', '3FD0C9', 'EDEFF2', 'C9CED6'
BAR = '63C384'
HEAD = ['Group', 'Name', 'Chargeable', 'Non-chargeable', 'Leave hours',
        'Sum of Public Holiday', 'Total Workable Hours', 'Total Hours incl of PH',
        'Chargeable %', 'Non-chargeable %', 'Chargeable Target',
        'Over / Under', '% of Leave']
WIDTHS = [13, 22, 12, 16, 12, 18, 18, 19, 13, 16, 16, 13, 11]
OVER = 12                                   # the Over / Under column
HOURS = (3, 4, 5, 6, 7, 8)                  # formatted #,##0
PCTS = (9, 10, 11, 12, 13)                  # one decimal, as 'Utilisation Q1' does

# Exactly as Danica sent it: chargeable, non-chargeable, leave, public holiday,
# total workable hours, chargeable target.
DATA = [
    ('Production', [
        ('Jaazaniah SALANOA',   574,  549, 190, 32, 1313, 0.65),
        ('Tyler WOOD',          957,  818, 319, 72, 2016, 0.65),
        ('Milo RANKIN',         792,  976, 284, 72, 2016, 0.65),
        ('Blake Crisford',      749,  785, 483, 72, 2016, 0.70),
        ('Sebastien Hathaway',  829, 1015, 173, 72, 2017, 0.70)]),
    ('Consulting', [
        ('Daniel SOBKOWSKI',    461, 1293, 261, 72, 2015, 0.50)]),
    ('Support', [
        ('Wade TONNA',         1843,    2, 214, 72, 2016, 0.88),
        ('Isaac BUTTERWORTH',  1757,   12, 174, 72, 1943, 0.88),
        ('Danielle HURLEY',    1846,    2, 172, 72, 2020, 0.88),
        ('Tyrell SHORTT',      1745,    0, 114, 48, 1859, 0.88),
        ('Tawonashe MADANGURE', 948,  191, 131, 36, 1269, 0.88),
        ('Kyle Krishnappa',    1772,    4, 241, 72, 2016, 0.88),
        ('Carlo Daru',         1803,    0, 205, 80, 2008, 0.88),
        ('Tina Nguyen',        1830,    0, 195, 72, 2016, 0.88),
        ('Kent Vincent Dejos',  931,    0,  86, 32, 1017, 0.88),
        ('Sangeeta Lachmaiya',  702,    0,  34,  8,  736, 0.88),
        ('Dylan Satherley',    1619,  168, 174, 72, 1960, 0.88),
        ('Nicholas Pinnington',1749,    8, 259, 72, 2016, 0.88),
        ('Jason Ng',           1839,    0, 182, 72, 2016, 0.88),
        ('Chloe Landayan',     1517,    0, 168, 32, 1685, 0.88),
        ('Layla Phillips',      665,    0,  32,  8,  697, 0.88),
        ('John Abbas Rizvi',   1336,    0,  40, 40, 1376, 0.88)]),
]


def derived(ws, r):
    """The five calculated columns, written against one row."""
    ws.cell(r, 8, '=G%d+F%d' % (r, r))
    ws.cell(r, 9, '=IFERROR(C%d/G%d,"")' % (r, r))
    ws.cell(r, 10, '=IFERROR(1-I%d,"")' % r)
    ws.cell(r, 12, '=IFERROR(I%d-K%d,"")' % (r, r))
    ws.cell(r, 13, '=IFERROR(E%d/G%d,"")' % (r, r))


def main():
    pk = openpyxl.load_workbook(PACK)
    # rebuilt from scratch each run, so a second run cannot leave two data
    # bars stacked on the same column or stale rows below the table
    if SHEET in pk.sheetnames:
        del pk[SHEET]
    ws = pk.create_sheet(SHEET, pk.sheetnames.index('Utilisation Q1') + 1)
    thin = Side(style='thin', color=LINE)

    ws.cell(TOP, 1).value = 'UTILISATION PER EMPLOYEE, ROLLING FORECAST'
    ws.cell(TOP, 1).font = Font(name='Calibri', bold=True, size=10, color=INK)
    head = TOP + 1
    for c, h in enumerate(HEAD, 1):
        cell = ws.cell(head, c, h)
        cell.font = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
        cell.fill = PatternFill('solid', fgColor=INK)
        cell.alignment = Alignment(horizontal='left' if c <= 2 else 'right',
                                   vertical='center', wrap_text=True)
        cell.border = Border(bottom=Side(style='medium', color=TEAL))
    ws.row_dimensions[head].height = 30

    def style(r, bold, band):
        for c in range(1, len(HEAD) + 1):
            cell = ws.cell(r, c)
            cell.font = Font(name='Calibri', bold=bold, size=9, color=INK)
            cell.alignment = Alignment(horizontal='left' if c <= 2 else 'right')
            if c in HOURS:
                cell.number_format = '#,##0'
            elif c in PCTS:
                cell.number_format = '0.0%'
            cell.border = Border(bottom=thin)
            if band:
                cell.fill = PatternFill('solid', fgColor=BAND)

    out = head + 1
    totals = []
    for group, people in DATA:
        first = out
        for i, (name, ch, nc, lv, ph, wk, tg) in enumerate(people):
            if i == 0:
                ws.cell(out, 1, group)
            ws.cell(out, 2, name)
            for c, v in zip((3, 4, 5, 6, 7), (ch, nc, lv, ph, wk)):
                ws.cell(out, c, v)
            ws.cell(out, 11, tg)
            derived(ws, out)
            style(out, False, False)
            out += 1
        last = out - 1
        ws.cell(out, 1, group + ' Total')
        for c in (3, 4, 5, 6, 7):
            ws.cell(out, c, '=SUM(%s%d:%s%d)' % (L(c), first, L(c), last))
        ws.cell(out, 11, '=IFERROR(SUMPRODUCT(G%d:G%d,K%d:K%d)/SUM(G%d:G%d),"")'
                % (first, last, first, last, first, last))
        derived(ws, out)
        style(out, True, True)
        totals.append(out)
        out += 1

    g = out
    ws.cell(g, 1, 'Grand Total')
    for c in (3, 4, 5, 6, 7):
        ws.cell(g, c, '=' + '+'.join('%s%d' % (L(c), t) for t in totals))
    ws.cell(g, 11, '=IFERROR(SUMPRODUCT(G%d:G%d,K%d:K%d)/SUM(G%d:G%d),"")'
            % (head + 1, g - 1, head + 1, g - 1, head + 1, g - 1))
    derived(ws, g)
    style(g, True, True)
    # the weighted target has to read the people rows only, not the subtotals
    peeps = [r for r in range(head + 1, g) if r not in totals]
    ws.cell(g, 11, '=IFERROR((%s)/(%s),"")' % (
        '+'.join('G%d*K%d' % (p, p) for p in peeps),
        '+'.join('G%d' % p for p in peeps)))

    for c, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[L(c)].width = w
    ws.freeze_panes = 'C%d' % (head + 1)

    rng = '%s%d:%s%d' % (L(OVER), head + 1, L(OVER), g)
    ws.conditional_formatting.add(rng, DataBarRule(
        start_type='min', end_type='max', color=BAR, showValue=True))

    pk.save(PACK)
    print('  schedule written to a new sheet %r, rows %d-%d' % (SHEET, head, g))
    print('  data bar %s on %s' % (BAR, rng))
    print("  snip range '%s'!A%d:M%d" % (SHEET, head, g))


if __name__ == '__main__':
    main()
