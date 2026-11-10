"""Put the departments into the year to date revenue chart.

The chart on Xero History rows 33 to 58 is the cumulative revenue line chart,
FY25, FY26, FY27 and the FY27 budget. Danica asked for the departments to be
in it.

Department lines will not work here. The year lines run to seven million and
the departments to Support's 715k, so five more lines would sit squashed along
the bottom. Stacked columns do work, because the five departments add to the
FY27 total exactly: 715,017 + 420,971 + 40,270 + 441,414 + 6,788 is 1,624,460,
which is the FY27 cumulative figure at September. So the stack tops out on the
FY27 line and shares the same axis.

The data is already in the pack. Rows 242 to 253 of Xero History hold
cumulative revenue by department for the twelve months, filled for Jul, Aug
and Sep. Nothing new is typed in, the chart just points at it.

The bar fills are the department brand colours lightened. Full strength would
collide with the lines: Support red is the budget line, Events teal is the
FY27 line and Video is the FY26 line. Lightening keeps each department its own
colour and leaves the lines reading on top.
"""
import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.fill import ColorChoice
from openpyxl.utils import get_column_letter

PACK = 'CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Xero History'
HEAD, FIRST, LAST = 241, 242, 253          # header row, first month, last month
DEPT_COLS = {2: 'F33844', 3: '005358', 4: '3FD0C9', 5: '6D71FF', 6: '8F4AFA'}
TINT = 0.30


def tint(hexstr, amount):
    r, g, b = (int(hexstr[i:i + 2], 16) for i in (0, 2, 4))
    f = lambda v: int(round(v + (255 - v) * amount))
    return '%02X%02X%02X' % (f(r), f(g), f(b))


def find(ws):
    """Return the anchored chart and the line part inside it.

    Once a bar chart has been combined in, openpyxl reads the pair back with
    the bar chart as the outer object and the lines as a part, so looking only
    at the outer object's own series finds nothing on a rerun.
    """
    for ch in ws._charts:
        for part in (list(getattr(ch, '_charts', [])) or [ch]):
            refs = [s.val.numRef.f for s in part.series if s.val and s.val.numRef]
            if any('$B$58' in (r or '') for r in refs):
                return ch, part
    return None, None


def main():
    wb = openpyxl.load_workbook(PACK)
    ws = wb[SHEET]

    holder, line = find(ws)
    if line is None:
        print('STOPPED: cumulative chart not found on %s' % SHEET)
        return
    for col in DEPT_COLS:
        if ws.cell(HEAD, col).value is None:
            print('STOPPED: no department name at row %d col %d' % (HEAD, col))
            return

    bars = BarChart()
    bars.type = 'col'
    bars.grouping = 'stacked'
    bars.overlap = 100
    bars.gapWidth = 60
    for col, colour in DEPT_COLS.items():
        ser = Series(Reference(ws, min_col=col, min_row=FIRST, max_row=LAST),
                     title_from_data=False)
        # the legend name has to be a StrRef to the header cell. Building it
        # through Series() leaves the name empty and the legend reads
        # "Column B" and so on.
        ser.tx = SeriesLabel(strRef=StrRef(
            "'%s'!$%s$%d" % (SHEET, get_column_letter(col), HEAD)))
        ser.graphicalProperties = GraphicalProperties(
            solidFill=ColorChoice(srgbClr=tint(colour, TINT)))
        ser.graphicalProperties.line.noFill = True
        bars.series.append(ser)
    bars.set_categories(Reference(ws, min_col=1, min_row=FIRST, max_row=LAST))

    # share the line chart's axes so the columns sit on the same scale
    bars.x_axis = line.x_axis
    bars.y_axis = line.y_axis

    # the parts list is what gets written. Bars first so the lines draw over
    # them, and rebuilding it from scratch drops any bars left by a rerun.
    holder._charts = [bars, line]

    wb.save(PACK)
    print('  department columns added, %d series over rows %d-%d'
          % (len(bars.series), FIRST, LAST))
    for col in DEPT_COLS:
        print('     %-12s %s' % (ws.cell(HEAD, col).value,
                                 tint(DEPT_COLS[col], TINT)))


if __name__ == '__main__':
    main()
