"""Rebuild the year to date chart from scratch, columns and lines together.

patch_v106 writes the reporting group block. This builds the chart.

Everything about the chart is created here: both plots, their series, their
colours and the axes. Nothing is reused from the chart that was there. Three
earlier attempts tried to add columns to the existing line chart, by editing
it in place, by editing the series after LibreOffice had rewritten it, and by
lifting a clean copy out of an older pack. All three came out with the four
year lines stacked on top of each other, because combining charts in openpyxl
carries the column plot's stacked grouping across unless both plots are built
in the same pass.

The lines keep the pack's own colours: FY25 navy, FY26 light teal, FY27 teal
and the budget red and dashed.
"""
import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.fill import ColorChoice
from openpyxl.utils import get_column_letter

PACK = 'CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Xero History'
L_HEAD, L_FIRST, L_LAST = 57, 58, 69       # the year lines
G_HEAD, G_FIRST, G_LAST = 257, 258, 269    # the reporting group columns
# column, colour, dashed, last row (FY27 only runs to September)
LINES = [(2, '122B82', False, 69), (3, '3FD0C9', False, 69),
         (4, '005358', False, 60), (5, 'F33844', True, 69)]
COLS = [(2, 'F33844'), (3, '005358'), (4, '6D71FF')]
TINT = 0.30


def tint(h, a):
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    f = lambda v: int(round(v + (255 - v) * a))
    return '%02X%02X%02X' % (f(r), f(g), f(b))


def label(col, row):
    return SeriesLabel(strRef=StrRef("'%s'!$%s$%d" % (SHEET, get_column_letter(col), row)))


def main():
    wb = openpyxl.load_workbook(PACK)
    ws = wb[SHEET]
    if ws.cell(G_HEAD, 2).value != 'Support':
        print('STOPPED: no reporting group block at row %d, run patch_v106 first'
              % G_HEAD)
        return
    old = [c for c in ws._charts
           if any('$B$58' in (s.val.numRef.f or '')
                  for p in (list(getattr(c, '_charts', [])) or [c])
                  for s in p.series if s.val and s.val.numRef)]
    if len(old) != 1:
        print('STOPPED: expected one year to date chart, found %d' % len(old))
        return

    bars = BarChart()
    bars.type, bars.grouping, bars.overlap, bars.gapWidth = 'col', 'stacked', 100, 60
    for col, colour in COLS:
        s = Series(Reference(ws, min_col=col, min_row=G_FIRST, max_row=G_LAST),
                   title_from_data=False)
        s.tx = label(col, G_HEAD)
        s.graphicalProperties = GraphicalProperties(
            solidFill=ColorChoice(srgbClr=tint(colour, TINT)))
        s.graphicalProperties.line.noFill = True
        bars.series.append(s)
    bars.set_categories(Reference(ws, min_col=1, min_row=G_FIRST, max_row=G_LAST))

    lines = LineChart()
    lines.grouping = 'standard'
    for col, colour, dashed, last in LINES:
        s = Series(Reference(ws, min_col=col, min_row=L_FIRST, max_row=last),
                   title_from_data=False)
        s.tx = label(col, L_HEAD)
        lp = LineProperties(solidFill=colour, w=28575)
        if dashed:
            lp.dashStyle = 'dash'
        s.graphicalProperties = GraphicalProperties(ln=lp)
        s.marker = Marker(symbol='none' if dashed else 'circle', size=6)
        s.smooth = False
        lines.series.append(s)
    lines.set_categories(Reference(ws, min_col=1, min_row=L_FIRST, max_row=L_LAST))

    # both plots are built here, then joined. The columns go in first so the
    # lines draw over them.
    bars += lines
    bars.y_axis.numFmt = '#,##0'
    bars.y_axis.delete = False
    bars.x_axis.delete = False
    bars.legend.position = 'b'
    bars.width, bars.height = 15, 7.5
    bars.anchor = old[0].anchor

    ws._charts.remove(old[0])
    ws._charts.append(bars)
    wb.save(PACK)
    print('  chart rebuilt: %d columns, %d lines'
          % (len(COLS), len(LINES)))


if __name__ == '__main__':
    main()
