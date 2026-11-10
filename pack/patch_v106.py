"""Slide 8's columns become the three reporting groups.

Danica asked for Production and Video combined, and Consulting and
Integration combined, on that slide only. That is the three groups the P&L is
built on: Support, Production (Events plus Video) and Consulting (Integration
plus Consulting).

A new block goes in at Xero History row 256 rather than changing the five
department block at row 241, because a second chart still reads that one and
the five department cut is wanted everywhere else. The new block adds the
existing columns up, it does not go back to source, so the three groups
cannot disagree with the five departments above them.

This has to run on a pack whose year to date chart is still a plain line
chart. Once LibreOffice has rewritten a combined chart, openpyxl can list it
but not reliably change it: editing the series left the four year lines
stacked on top of each other, and so did lifting a clean chart out of an
older copy of the pack and dropping it in. Rebuilding the pack from the last
commit before the columns went in, and combining in place, is what works.
"""
import openpyxl
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.fill import ColorChoice
from openpyxl.utils import get_column_letter

PACK = 'CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Xero History'
SRC_FIRST, SRC_LAST = 242, 253             # the five department block
TOP = 256
HEAD, FIRST, LAST = TOP + 1, TOP + 2, TOP + 13
# name, the sum over the source row, the cell that says whether that month is
# filled in yet, and the brand colour
GROUPS = [('Support',    '$B%(r)d',         'B', 'F33844'),
          ('Production', '$C%(r)d+$D%(r)d', 'C', '005358'),
          ('Consulting', '$E%(r)d+$F%(r)d', 'E', '6D71FF')]
TINT = 0.30


def tint(h, a):
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    f = lambda v: int(round(v + (255 - v) * a))
    return '%02X%02X%02X' % (f(r), f(g), f(b))


def main():
    wb = openpyxl.load_workbook(PACK)
    ws = wb[SHEET]
    line = None
    for ch in ws._charts:
        refs = [s.val.numRef.f for s in ch.series if s.val and s.val.numRef]
        if any('$B$58' in (r or '') for r in refs):
            line = ch
            break
    if line is None:
        print('STOPPED: no plain year to date line chart. Rebuild the pack from'
              ' the last commit before the columns went in and run this again.')
        return
    if len(line._charts) != 1:
        print('STOPPED: the chart has already been combined, %d parts'
              % len(line._charts))
        return
    if ws.cell(SRC_FIRST, 2).value is None:
        print('STOPPED: no department block at row %d' % SRC_FIRST)
        return

    ws.cell(TOP, 1).value = 'YEAR TO DATE REVENUE BY REPORTING GROUP, CUMULATIVE'
    ws.cell(HEAD, 1).value = 'Month'
    for n, (name, _, _, _) in enumerate(GROUPS):
        ws.cell(HEAD, 2 + n).value = name
    for i, src in enumerate(range(SRC_FIRST, SRC_LAST + 1)):
        r = FIRST + i
        ws.cell(r, 1).value = '=A%d' % src
        for n, (name, expr, guard, _) in enumerate(GROUPS):
            # an empty source month stays empty, or the chart draws a zero
            # column for every month still to come
            ws.cell(r, 2 + n).value = ('=IF($%s%d="","",%s)'
                                       % (guard, src, expr % {'r': src}))

    bars = BarChart()
    bars.type, bars.grouping, bars.overlap, bars.gapWidth = 'col', 'stacked', 100, 60
    for n, (name, _, _, colour) in enumerate(GROUPS):
        col = 2 + n
        ser = Series(Reference(ws, min_col=col, min_row=FIRST, max_row=LAST),
                     title_from_data=False)
        ser.tx = SeriesLabel(strRef=StrRef(
            "'%s'!$%s$%d" % (SHEET, get_column_letter(col), HEAD)))
        ser.graphicalProperties = GraphicalProperties(
            solidFill=ColorChoice(srgbClr=tint(colour, TINT)))
        ser.graphicalProperties.line.noFill = True
        bars.series.append(ser)
    bars.set_categories(Reference(ws, min_col=1, min_row=FIRST, max_row=LAST))
    bars.x_axis, bars.y_axis = line.x_axis, line.y_axis
    # columns first in the plot area so the lines draw over them
    line._charts = [bars] + list(line._charts)

    wb.save(PACK)
    print('  group block at rows %d-%d' % (FIRST, LAST))
    for name, _, _, colour in GROUPS:
        print('     %-12s %s' % (name, tint(colour, TINT)))


if __name__ == '__main__':
    main()
