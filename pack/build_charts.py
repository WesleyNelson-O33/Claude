"""Every chart in the pack, built in one place with the same settings.

Clears whatever charts are on the sheets and puts them back, because openpyxl
does not round trip data label settings: reload a book and the labels come
back as None, so patching them in place is not reliable. Rebuilding is.

House rules for a chart here:
  data labels show the value and nothing else, never the series name, the
    category name or a legend key
  a ranked bar chart runs best at the top, and drops the value axis, since
    the label on each bar already says the number
  bars are thick, so the chart reads as a shape rather than a row of pins
  colours come from brand.json, and departments from dept_colours.json

Run after the patch scripts and before recalc.
"""
import os, json
import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.legend import Legend
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (RichTextProperties, Paragraph,
                                   ParagraphProperties, CharacterProperties)
from openpyxl.utils import get_column_letter as gl

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = f'{HERE}/CTS Financial Controller Pack - v20.xlsx'
B = json.load(open(f'{HERE}/data/brand.json'))
D = json.load(open(f'{HERE}/data/dept_colours.json'))
P = json.load(open(f'{HERE}/data/pldept.json'))
R = json.load(open(f'{HERE}/data/newsheet_rows.json'))
ACC = json.load(open(f'{HERE}/data/overhead_accounts.json'))
CC = json.load(open(f'{HERE}/data/costcharts.json'))

INK, GREY, GREYL, WHITE = B['charcoal'], B['grey'], B['grey_lt'], B['white']
TEAL, AQUA, NAVY, BLUE, PURP, RED = (B['dark_teal'], B['bright_teal'], B['deep_blue'],
                                     B['primary_blue'], B['purple'], B['accent_red'])

def txt(size=900, bold=False, colour=INK):
    return RichText(bodyPr=RichTextProperties(), p=[Paragraph(
        pPr=ParagraphProperties(defRPr=CharacterProperties(
            sz=size, b=bold, solidFill=colour, latin=None)),
        endParaRPr=CharacterProperties(sz=size, b=bold, solidFill=colour))])

def _dl(fmt, pos, size, colour, bold):
    d = DataLabelList()
    d.showVal = True
    d.showSerName = False
    d.showCatName = False
    d.showLegendKey = False
    d.showPercent = False
    d.showBubbleSize = False
    d.numFmt = fmt
    d.dLblPos = pos
    d.txPr = txt(size, bold, colour)
    return d

def labels(ch, fmt='#,##0', pos='outEnd', size=900, colour=INK, bold=False):
    ch.dataLabels = _dl(fmt, pos, size, colour, bold)

def labels_out(ch, fmt='"$"#,##0', size=900):
    """Value outside the bar end, in charcoal. Use where the figures have a
    long tail: a centred label does not fit inside a short bar, and these
    lists run from tens of thousands down to hundreds."""
    ch.dataLabels = _dl(fmt, 'outEnd', size, INK, False)

def ranked_axis(ch):
    """Ranked, but the value axis stays. It gives the longest bar's label room
    at the right hand end, which outEnd labels need."""
    ch.x_axis.scaling.orientation = 'maxMin'
    ch.x_axis.tickLblPos = 'low'
    ch.y_axis.majorGridlines = None

def hide_thin(ch, shares, floor=0.045):
    """Drop the label on any segment too narrow to hold it.

    A hundred per cent stacked bar can only put the label inside the segment.
    Where a segment is half a per cent of the bar, its number prints across the
    neighbour and both become unreadable. The figures are all on the table
    beside the chart, so dropping the label loses nothing."""
    from openpyxl.chart.label import DataLabel
    off = dict(showVal=False, showSerName=False, showCatName=False,
               showLegendKey=False, showPercent=False, showBubbleSize=False)
    for s_, row in zip(ch.series, shares):
        thin = [DataLabel(idx=i, **off)
                for i, v in enumerate(row) if v < floor]
        if thin:
            s_.dLbls.dLbl = thin


def mix_shares(wb, bases, n=3, groups=None):
    """Each segment as a share of its own bar, for the reporting month. Read
    from the pasted hours, not from the chart cells, because those are formulas
    and an openpyxl save leaves them with no cached value."""
    u, want = wb['Utilisation'], wb['Summary']['B3'].value
    col = next(c for c in range(4, 16) if u.cell(4, c).value == want)
    rows = [[u.cell(b + i, col).value or 0 for b in bases] for i in range(n)]
    if groups:
        rows = [[sum(r[k] for k in g) for g in groups] for r in rows]
    w = len(rows[0])
    tot = [sum(rows[i][j] for i in range(n)) or 1 for j in range(w)]
    return [[rows[i][j] / tot[j] for j in range(w)] for i in range(n)]


def labels_in(ch, cols, fmt='$#,##0', size=900):
    """Sit the value in the middle of each bar. At the end of the bar the
    longest one runs into the edge of the plot and its label gets dropped.
    One text colour per series, because a white number is invisible on a
    pale bar."""
    for s_, col in zip(ch.series, cols):
        s_.dLbls = _dl(fmt, 'ctr', size, col, True)

def frame(ch, title, w, h, gap=45):
    ch.title = title
    ch.width, ch.height = w, h
    ch.style = None
    ch.gapWidth = gap
    ch.y_axis.majorGridlines = None
    ch.legend = None
    for ax in (ch.x_axis, ch.y_axis):
        ax.txPr = txt(900)
        ax.majorTickMark = 'none'
        ax.minorTickMark = 'none'
        ax.spPr = GraphicalProperties(ln=LineProperties(solidFill=GREY, w=9525))

def ranked(ch):
    """A top ten reads best with number one at the top and no value axis.
    tickLblPos low pins the names to the left; without it they drift to
    wherever the value axis crosses and end up sitting on the bars."""
    ch.x_axis.scaling.orientation = 'maxMin'
    ch.x_axis.tickLblPos = 'low'
    ch.y_axis.delete = True
    ch.y_axis.majorGridlines = None

def legend(ch, pos='b'):
    ch.legend = Legend()
    ch.legend.position = pos
    ch.legend.overlay = False
    ch.legend.txPr = txt(900)

def paint(ch, cols):
    for s, col in zip(ch.series, cols):
        s.graphicalProperties = GraphicalProperties(solidFill=col)
        s.graphicalProperties.line = LineProperties(noFill=True)

def paint_depts(ser, names):
    ser.data_points = [DataPoint(idx=i, spPr=GraphicalProperties(
        solidFill=D['departments'].get(n, INK), ln=LineProperties(noFill=True)))
        for i, n in enumerate(names)]

wb = openpyxl.load_workbook(OUT)
for ws in wb.worksheets:
    ws._charts = []

# ---------------------------------------------------- Top 10 Clients ------
ws = wb['Top 10 Clients']
ch = BarChart(); ch.type = 'bar'; ch.grouping = 'clustered'
ch.add_data(Reference(ws, min_col=4, min_row=6, max_row=16), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=7, min_row=7, max_row=16))
frame(ch, 'Top 10 customers, year to date', 24, 12)
ranked(ch)
paint(ch, [TEAL])
# The story on this chart is concentration, so the top three are called out in
# the Support red and the rest left in teal.
ch.series[0].data_points = [
    DataPoint(idx=i, spPr=GraphicalProperties(
        solidFill=RED if i < 3 else TEAL, ln=LineProperties(noFill=True)))
    for i in range(10)]
labels_in(ch, [WHITE])
ws.add_chart(ch, 'G5')

# -------------------------------------------------------- Top 10 YoY ------
ws = wb['Top 10 YoY']
ch = BarChart(); ch.type = 'bar'; ch.grouping = 'clustered'
ch.add_data(Reference(ws, min_col=3, max_col=4, min_row=4, max_row=14),
            titles_from_data=True)
ch.set_categories(Reference(ws, min_col=10, min_row=5, max_row=14))
frame(ch, 'This year against the same three months last year', 24, 13, gap=55)
ranked(ch)
legend(ch)
paint(ch, [NAVY, GREY])
labels_in(ch, [WHITE, INK], size=800)
ws.add_chart(ch, 'J4')

# ------------------------------------------------------- P&L by Dept ------
# The six blocks are a P&L each, so the chart reads the side by side table at
# the bottom rather than any one block.
ws = wb['P&L by Dept']
S = P["side_by_side"]
N = 4                                       # four departments now
DNAMES = [ws.cell(S + 2 + i, 1).value for i in range(N)]
for cols, anchor, title in (((2, 4), f'I{S}', 'Revenue and gross profit by department, year to date'),
                            ((2, 6), f'I{S + 24}', 'Revenue and net profit by department, year to date')):
    ch = BarChart(); ch.type = 'col'; ch.grouping = 'clustered'
    for c in cols:
        ch.add_data(Reference(ws, min_col=c, min_row=S + 1, max_row=S + 1 + N),
                    titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=S + 2, max_row=S + 1 + N))
    frame(ch, title, 23, 11)
    labels(ch, '#,##0', size=850)
    legend(ch)
    paint(ch, [TEAL, INK])
    paint_depts(ch.series[0], DNAMES)
    ws.add_chart(ch, anchor)

# --------------------------------------------------------------- DSO ------
ws = wb['DSO']
ch = BarChart(); ch.type = 'col'
ch.add_data(Reference(ws, min_col=2, min_row=18, max_row=24), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=19, max_row=24))
frame(ch, 'Aged receivables', 18, 10)
labels(ch, '$#,##0')
ch.y_axis.delete = True
paint(ch, [NAVY])
ws.add_chart(ch, 'J5')

ch = BarChart(); ch.type = 'bar'
ch.add_data(Reference(ws, min_col=3, min_row=30, max_row=40), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, min_row=31, max_row=40))
frame(ch, 'Who owes the most', 18, 12)
ranked(ch)
paint(ch, [BLUE])
labels_in(ch, [WHITE])
ws.add_chart(ch, 'J24')

T = R['dso_trend']
ch = LineChart()
ch.add_data(Reference(ws, min_col=1, min_row=T + 5, max_col=13, max_row=T + 5),
            from_rows=True, titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, min_row=T + 1, max_col=13))
frame(ch, 'DSO month by month', 23, 10)
labels(ch, '#,##0.0', pos='t')
for s in ch.series:
    s.graphicalProperties = GraphicalProperties(
        ln=LineProperties(solidFill=TEAL, w=28000))
    s.smooth = False
ws.add_chart(ch, f'A{T + 9}')

# ---------------------------------------------------- Overhead Split ------
ws = wb['Overhead Split']
ch = BarChart(); ch.type = 'col'; ch.grouping = 'clustered'
ch.add_data(Reference(ws, min_col=1, min_row=16, max_col=6, max_row=18),
            from_rows=True, titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, min_row=15, max_col=6))
frame(ch, 'The three split keys', 21, 10)
labels(ch, '0%', size=850)
ch.y_axis.delete = True
legend(ch)
paint(ch, [INK, GREY, NAVY])
ws.add_chart(ch, 'V5')

# ---------------------------------------------------- Overheads MoM ------
ws = wb['Overheads MoM']
SUB = [r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value in ACC]
P0 = next(r for r in range(1, ws.max_row + 1)
          if str(ws.cell(r, 1).value or '').startswith('CHART DATA'))
first, last = P0 + 2, P0 + 1 + len(SUB)
ch = BarChart(); ch.type = 'col'; ch.grouping = 'stacked'; ch.overlap = 100
ch.add_data(Reference(ws, min_col=1, min_row=first, max_col=13, max_row=last),
            from_rows=True, titles_from_data=True)
ch.set_categories(Reference(ws, min_col=2, min_row=P0 + 1, max_col=13))
frame(ch, 'Overhead spend, month by month', 30, 13, gap=35)
legend(ch)
paint(ch, [NAVY, BLUE, PURP, INK, GREY, AQUA])
ws.add_chart(ch, f'A{last + 3}')

# -------------------------------------------- Top 10 by department -------
ws = wb['Top 10 Clients']
S = CC['top10_dept']
ch = BarChart(); ch.type = 'bar'; ch.grouping = 'stacked'; ch.overlap = 100
ch.add_data(Reference(ws, min_col=2, max_col=7, min_row=S + 1, max_row=S + 11),
            titles_from_data=True)
ch.set_categories(Reference(ws, min_col=7, min_row=7, max_row=16))
frame(ch, 'Top 10 customers, split by department', 23, 13, gap=45)
ranked(ch)
legend(ch)
paint(ch, [D['departments'][d] for d in
           ('Support', 'Events', 'Video', 'Integration', 'Consulting', 'General Admin')])
ws.add_chart(ch, f'J{S}')

# ------------------------------------------------- Cost Analysis ---------
# All seven sit in column I and down, stacked with a gap, so the data in
# columns A to G stays clear of them.
ws = wb['Cost Analysis']

# the operating expense accounts, biggest first
n = CC['ops_n']
ch = BarChart(); ch.type = 'bar'; ch.grouping = 'clustered'
ch.add_data(Reference(ws, min_col=2, min_row=CC['ops_first'] - 1,
                      max_row=CC['ops_first'] - 1 + n), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=CC['ops_first'],
                            max_row=CC['ops_first'] - 1 + n))
frame(ch, 'Other operating expenses, year to date', 22, 10, gap=40)
ranked_axis(ch)
paint(ch, [NAVY])
labels_out(ch)
ws.add_chart(ch, 'I4')

# where every dollar of operating cost goes
ch = BarChart(); ch.type = 'bar'; ch.grouping = 'clustered'
ch.add_data(Reference(ws, min_col=2, min_row=CC['dist_first'] - 1,
                      max_row=CC['dist_last']), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=CC['dist_first'],
                            max_row=CC['dist_last']))
frame(ch, 'Operating cost distribution, year to date', 22, 11, gap=40)
ranked_axis(ch)
paint(ch, [INK])
labels_out(ch)
ws.add_chart(ch, 'I33')

# direct against indirect labour
ch = BarChart(); ch.type = 'bar'; ch.grouping = 'stacked'; ch.overlap = 100
ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=CC['lab_first'] - 1,
                      max_row=CC['lab_last']), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=CC['lab_first'],
                            max_row=CC['lab_last']))
frame(ch, 'Labour mix, direct against indirect, year to date', 22, 13, gap=40)
ranked_axis(ch)
legend(ch)
paint(ch, [NAVY, GREY])
labels_out(ch, size=800)
ws.add_chart(ch, 'I61')

# ------------------------------------------- budget against actual ------
ws = wb['Cost Analysis']
ch = BarChart(); ch.type = 'bar'; ch.grouping = 'clustered'
ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=CC['bva_hdr'],
                      max_row=CC['bva_last']), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=CC['bva_first'],
                            max_row=CC['bva_last']))
frame(ch, 'Budget against actual, year to date', 22, 11, gap=50)
ranked_axis(ch)
legend(ch)
paint(ch, [NAVY, GREY])
labels_out(ch, size=800)
ws.add_chart(ch, 'I92')

ch = BarChart(); ch.type = 'col'; ch.grouping = 'clustered'
# Column D is the actual for closed months only. Column B carries the budget
# once the month is in the future, so charting B printed every label twice.
ch.add_data(Reference(ws, min_col=4, min_row=CC['bvm_hdr'], max_row=CC['bvm_last']),
            titles_from_data=True)
ch.add_data(Reference(ws, min_col=3, min_row=CC['bvm_hdr'], max_row=CC['bvm_last']),
            titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=CC['bvm_first'],
                            max_row=CC['bvm_last']))
frame(ch, 'Income by month, actual against budget', 26, 12, gap=60)
legend(ch)
paint(ch, [NAVY, GREY])
# Thousands, so twenty four labels fit side by side instead of colliding, and
# the trailing semicolons hide the zero on the months with no actual yet.
labels(ch, '#,##0,;;', size=800)
ws.add_chart(ch, 'I117')

# --------------------------------------------------- profitability ------
ws = wb['Cost Analysis']
ch = BarChart(); ch.type = 'col'; ch.grouping = 'clustered'
ch.add_data(Reference(ws, min_col=2, max_col=4, min_row=CC['prof_hdr'],
                      max_row=CC['prof_last']), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=CC['prof_first'],
                            max_row=CC['prof_last']))
frame(ch, 'Profitability month on month, actual to date then budget', 26, 11)
ch.y_axis.delete = True
legend(ch)
paint(ch, [NAVY, BLUE, INK])
labels(ch, '#,##0', size=700)
ws.add_chart(ch, 'I142')

ch = BarChart(); ch.type = 'col'; ch.grouping = 'clustered'
for col in (4, 5):
    ch.add_data(Reference(ws, min_col=col, min_row=CC['prof_hdr'],
                          max_row=CC['prof_last']), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=CC['prof_first'],
                            max_row=CC['prof_last']))
frame(ch, 'Net profit year on year, this year against last', 26, 11)
ch.y_axis.delete = True
legend(ch)
paint(ch, [NAVY, GREY])
labels(ch, '#,##0', size=700)
ws.add_chart(ch, 'I167')

# ------------------------------- revenue by department, the month -------
# Actual against budget and against the same month last year.
ws = wb['Cost Analysis']
ch = BarChart(); ch.type = 'col'; ch.grouping = 'clustered'
ch.add_data(Reference(ws, min_col=2, max_col=4, min_row=191, max_row=196),
            titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=192, max_row=196))
frame(ch, 'Revenue by department, the month against budget and last year', 26, 11)
ch.y_axis.delete = True
legend(ch)
paint(ch, [TEAL, GREY, AQUA])
labels(ch, '#,##0', size=700)
ws.add_chart(ch, 'F190')

# ----------------------------------- hours by department, the month ----
ws = wb['Utilisation']
# Support works roughly a hundred times the hours Integration does, so on an
# absolute scale four of the five departments are invisible. Each bar is scaled
# to its own total instead, which is what makes the mix readable. The label on
# each segment still carries the hours.
ch = BarChart(); ch.type = 'bar'; ch.grouping = 'percentStacked'; ch.overlap = 100
ch.add_data(Reference(ws, min_col=2, max_col=4, min_row=67, max_row=72),
            titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=68, max_row=72))
frame(ch, 'How each department spent its hours, billable against not billable',
      24, 13)
ranked(ch)
legend(ch)
paint(ch, [TEAL, AQUA, GREYL])
# A stacked bar can only take a label inside the segment, so the label colour
# has to follow the segment under it. White is invisible on the bright teal and
# on the pale grey.
labels_in(ch, [WHITE, INK, INK], fmt='#,##0', size=800)
hide_thin(ch, mix_shares(wb, [9, 15, 21, 27, 33]))
ws.add_chart(ch, 'F66')

# ------------------------- employment expense, office against dept -----
ws = wb['Overheads MoM']
ch = BarChart(); ch.type = 'bar'; ch.grouping = 'stacked'; ch.overlap = 100
ch.add_data(Reference(ws, min_col=2, max_col=3, min_row=161, max_row=166),
            titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=162, max_row=166))
frame(ch, 'Employment expense, departmental staff against office staff', 24, 10)
ranked(ch)
legend(ch)
paint(ch, [TEAL, AQUA])
# Same again: the office staff segment is bright teal, so its label is ink.
labels_in(ch, [WHITE, INK], fmt='$#,##0', size=700)
ws.add_chart(ch, 'E160')

# ------------------------- revenue by month and department --------------
# The mix behind each month's revenue, which the other revenue charts do not
# show: one column a month, split by the department that billed it.
ws = wb['Cost Analysis']
# only the closed months: an empty column for every month still to come makes
# the chart read as unfinished. The period number comes off the reporting
# month, so the chart lengthens on its own as the year goes on.
period = ((wb['Summary']['B3'].value.month - int(wb['Setup']['B6'].value)) % 12) + 1
ch = BarChart(); ch.type = 'col'; ch.grouping = 'stacked'; ch.overlap = 100
ch.add_data(Reference(ws, min_col=2, max_col=6, min_row=231,
                      max_row=231 + period), titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=232, max_row=231 + period))
frame(ch, 'Revenue by month and department, against budget', 30, 11, gap=60)
legend(ch)
# each department in its own colour, the same ones the introduction cards use
paint(ch, [D['departments'][k] for k in
                ('Support', 'Events', 'Video', 'Integration', 'Consulting')])
# no segment labels: five stacked segments a column collide at this size, and
# the figures behind them are on the P&L summary table

# budget for the same months, as a line over the stack, so the chart answers
# both questions at once: who billed it, and was the month enough
line = LineChart()
line.add_data(Reference(ws, min_col=7, min_row=231, max_row=231 + period),
              titles_from_data=True)
line.series[0].graphicalProperties = GraphicalProperties(
    ln=LineProperties(solidFill=INK, w=22000))
line.series[0].marker = Marker(symbol='circle', size=7)
line.series[0].marker.graphicalProperties = GraphicalProperties(solidFill=INK)
line.series[0].smooth = False
ch += line
ws.add_chart(ch, 'H230')

# ----------------- hours by the three departments, the month ------------
# The same picture as the five department version, rolled up the way the P&L
# groups the business: Production is Events and Video, Consulting is
# Integration and Consulting.
ws = wb['Utilisation']
ch = BarChart(); ch.type = 'bar'; ch.grouping = 'percentStacked'; ch.overlap = 100
ch.add_data(Reference(ws, min_col=2, max_col=4, min_row=77, max_row=80),
            titles_from_data=True)
ch.set_categories(Reference(ws, min_col=1, min_row=78, max_row=80))
frame(ch, 'How the three departments spent their hours', 24, 11)
ranked(ch)
legend(ch)
paint(ch, [TEAL, AQUA, GREYL])
labels_in(ch, [WHITE, INK, INK], fmt='#,##0', size=800)
hide_thin(ch, mix_shares(wb, [9, 15, 21, 27, 33], groups=[[0], [1, 2], [3, 4]]))
ws.add_chart(ch, 'F92')

wb.save(OUT)
print('charts rebuilt:', sum(len(ws._charts) for ws in wb.worksheets))
