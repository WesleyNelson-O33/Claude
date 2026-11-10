"""Top ten customers across three first quarters, into the pack.

Danica asked for the customer slide to compare three years. The pack only
held two, FY27 and FY26 year to date by customer on the Top 10 YoY sheet.

Q1 FY25, that is July to September 2024, was not in the pack at all. It is
summed here from the 187 Xero invoices for that quarter, net of GST. The
basis was checked against Xero's own revenue report for the same quarter:
every one of its top five comes to exactly 1.10 times the net sum, so that
report is GST inclusive and these net figures are the ones that match how the
pack counts a customer. Those five tied to the dollar.

The Q1 2024 column is a paste, like the other pasted blocks in the pack. It
does not recalculate. If the quarter is ever restated it has to be pulled
again.

It gets a sheet of its own so its column widths cannot squeeze the tables and
charts already snipped off other sheets.

Two blocks are written. The table reads biggest first. The chart needs the
opposite order to plot top down, so it reads off a mirror block built with
formulas pointing at the table, which cannot drift from it.
"""
import openpyxl
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.fill import ColorChoice
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

# Every Font here names Calibri. Without a name openpyxl leaves the
# workbook default, which comes out of LibreOffice as a serif and does
# not match the rest of the deck.
PACK = 'CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Top 10 Q1 3yr'
AFTER = 'Top 10 YoY'
INK, TEAL, BAND, LINE = '242C39', '3FD0C9', 'EDEFF2', 'C9CED6'
GOOD, BAD = '005358', 'C41A28'
# oldest faintest, newest strongest
YEARS = [('Q1 FY25', 'C9CED6'), ('Q1 FY26', '3FD0C9'), ('Q1 FY27', '122B82')]
HEAD = ['Customer', 'Q1 FY25', 'Q1 FY26', 'Q1 FY27', 'Change on FY26', 'Note']
WIDTHS = [24, 13, 13, 13, 17, 17]
# name, Q1 2024, Q1 2025, Q1 2026, note. Ordered by Q1 2026, as the slide is.
ROWS = [
    ('Deloitte Services', 357134, 381989, 401755, ''),
    ('Aware Super',            0,      0, 285075, 'new this year'),
    ('CBA Place South',   174031, 119630, 105040, 'down three years'),
    ('Ventia',             30030,  59169,  97662, 'up three years'),
    ('PwC Services',       54926,  58325,  88051, 'up three years'),
    ('APT Management',     73233,   5669,  63613, ''),
    ('ASX Operations',     91795,  89996,  60858, 'down three years'),
    ('Bankwest',           49243,  60854,  54971, ''),
    ('Data#3',                 0,      0,  45909, 'new this year'),
    ('Generation-e',           0,      0,  37166, 'new this year'),
]
TOP = 1
HEAD_ROW = TOP + 1
FIRST = HEAD_ROW + 1
LAST = FIRST + len(ROWS) - 1
MIRROR = LAST + 4                        # header of the block the chart reads


def main():
    wb = openpyxl.load_workbook(PACK)
    if SHEET in wb.sheetnames:
        del wb[SHEET]
    ws = wb.create_sheet(SHEET, wb.sheetnames.index(AFTER) + 1)
    thin = Side(style='thin', color=LINE)

    ws.cell(TOP, 1).value = 'TOP TEN CUSTOMERS, FIRST QUARTER, THREE YEARS'
    ws.cell(TOP, 1).font = Font(name='Calibri', bold=True, size=10, color=INK)
    for c, h in enumerate(HEAD, 1):
        cell = ws.cell(HEAD_ROW, c, h)
        cell.font = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
        cell.fill = PatternFill('solid', fgColor=INK)
        cell.alignment = Alignment(
            horizontal='left' if c in (1, 6) else 'right',
            indent=1 if c == 6 else 0, vertical='center', wrap_text=True)
        cell.border = Border(bottom=Side(style='medium', color=TEAL))
    ws.row_dimensions[HEAD_ROW].height = 20

    for i, (name, a, b, c3, note) in enumerate(ROWS):
        r = FIRST + i
        ws.cell(r, 1, name).font = Font(name='Calibri', size=9, color=INK)
        ws.cell(r, 1).border = Border(bottom=thin)
        for j, v in enumerate((a, b, c3)):
            cell = ws.cell(r, 2 + j, v)
            cell.font = Font(name='Calibri', size=9, color=INK)
            cell.number_format = '#,##0'
            cell.alignment = Alignment(horizontal='right')
            cell.border = Border(bottom=thin)
        chg = ws.cell(r, 5, '=D%d-C%d' % (r, r))
        chg.number_format = '+#,##0;-#,##0;-'
        chg.font = Font(name='Calibri', size=9, bold=True, color=GOOD if c3 >= b else BAD)
        chg.alignment = Alignment(horizontal='right')
        chg.border = Border(bottom=thin)
        n = ws.cell(r, 6, note)
        n.font = Font(name='Calibri', size=9, italic=True, color=INK)
        n.alignment = Alignment(horizontal='left', indent=1)
        n.border = Border(bottom=thin)

    tot = LAST + 1
    ws.cell(tot, 1, 'These ten together').font = Font(name='Calibri', bold=True, size=9, color=INK)
    for j, col in enumerate('BCD'):
        cell = ws.cell(tot, 2 + j, '=SUM(%s%d:%s%d)' % (col, FIRST, col, LAST))
        cell.font = Font(name='Calibri', bold=True, size=9, color=INK)
        cell.number_format = '#,##0'
    ws.cell(tot, 5, '=D%d-C%d' % (tot, tot)).number_format = '+#,##0;-#,##0;-'
    ws.cell(tot, 5).font = Font(name='Calibri', bold=True, size=9, color=GOOD)
    for c in range(1, 7):
        ws.cell(tot, c).fill = PatternFill('solid', fgColor=BAND)

    note = tot + 1
    ws.cell(note, 1).value = (
        'The ten are this quarter\u2019s top ten. The earlier columns are what those '
        'same ten were worth then, not the top ten of those years.')
    ws.cell(note + 1, 1).value = (
        'Q1 FY25 is pasted: the 187 Xero invoices for 1 Jul to 30 Sep 2024, net of '
        'GST. It does not recalculate.')
    ws.cell(note + 1, 1).font = Font(name='Calibri', size=8, italic=True, color='6B7585')
    ws.cell(note, 1).font = Font(name='Calibri', size=8, italic=True, color='6B7585')

    # the mirror the chart reads, smallest first so it plots top down
    for c, h in enumerate(HEAD[:4], 1):
        ws.cell(MIRROR, c, h).font = Font(name='Calibri', bold=True, size=9, color=INK)
    for i in range(len(ROWS)):
        src = LAST - i
        r = MIRROR + 1 + i
        for c in range(1, 5):
            ws.cell(r, c, '=%s%d' % (L(c), src))
            ws.cell(r, c).number_format = '#,##0' if c > 1 else 'General'
            ws.cell(r, c).font = Font(name='Calibri', size=9, color=INK)

    ch = BarChart()
    ch.type, ch.grouping, ch.gapWidth, ch.overlap = 'bar', 'clustered', 40, -10
    for j, (label, colour) in enumerate(YEARS):
        col = 2 + j
        s = Series(Reference(ws, min_col=col, min_row=MIRROR + 1,
                             max_row=MIRROR + len(ROWS)), title_from_data=False)
        s.tx = SeriesLabel(strRef=StrRef("'%s'!$%s$%d" % (SHEET, L(col), MIRROR)))
        s.graphicalProperties = GraphicalProperties(solidFill=ColorChoice(srgbClr=colour))
        s.graphicalProperties.line.noFill = True
        ch.series.append(s)
    ch.set_categories(Reference(ws, min_col=1, min_row=MIRROR + 1,
                                max_row=MIRROR + len(ROWS)))
    ch.legend.position = 'b'
    ch.y_axis.numFmt = '#,##0'
    ch.x_axis.delete = ch.y_axis.delete = False
    ch.width, ch.height = 30, 10.2
    ws.add_chart(ch, 'A%d' % (MIRROR + len(ROWS) + 3))

    for c, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[L(c)].width = w

    wb.save(PACK)
    print('  sheet %r written' % SHEET)
    print('  table  A%d:F%d' % (HEAD_ROW, tot))
    print('  chart  A%d:V%d' % (MIRROR + len(ROWS) + 3, MIRROR + len(ROWS) + 29))


if __name__ == '__main__':
    main()
