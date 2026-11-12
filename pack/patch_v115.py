"""Add the rolling forecast version of the hours by department chart.

Slide 27 shows billable, non-billable and leave as a stacked bar per
department for the quarter. Danica asked for the same thing on the rolling
forecast, so this builds it on the 'Utilisation Rolling' sheet and the snip
comes off the pack like the rest.

Three bars, not five. Her rolling pivot groups people by the team they sit
in, which is Support, Production and Consulting. The five way split on
slide 27 is a different cut, the department the hours were booked to, and
it is only held for the months already closed. There is no rolling forecast
version of it, so this one cannot be built five ways without that data.

The chart copies chart16 on the Utilisation sheet, which is the three
department version of the same picture: horizontal bars, percent stacked,
overlap 100, gap width 45, legend at the bottom, no chart title, labels in
#,##0. Colours are the pack's own, 005358 billable, 3fd0c9 non-billable,
f4f4f4 leave.

Run it with no arguments.
"""
# Every Font here names Calibri, or LibreOffice renders the default as serif.
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import (CharacterProperties, Font as DFont, Paragraph,
                                   ParagraphProperties, RichTextProperties)
from openpyxl.drawing.fill import PatternFillProperties
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

PACK = 'CTS Financial Controller Pack - v20.xlsx'
SHEET = 'Utilisation Rolling'
INK, TEAL, BAND, LINE = '242C39', '3FD0C9', 'EDEFF2', 'C9CED6'
BILL, NONBILL, LEAVE = '005358', '3fd0c9', 'f4f4f4'
TOP = 31                                  # the block sits under the table
# group total rows written by patch_v113
ROWS = [('Support', 27), ('Production', 8), ('Consulting', 10)]
CH, NC, LV = 'C', 'D', 'E'                # chargeable, non-chargeable, leave


def labels(colour):
    """Data labels in her style: bold, 8pt, Calibri, #,##0."""
    pr = CharacterProperties(b=True, sz=800, latin=DFont(typeface='Calibri'),
                             solidFill=colour)
    dl = DataLabelList(showVal=True, showSerName=False, showCatName=False,
                       showLegendKey=False, showPercent=False, showBubbleSize=False)
    dl.numFmt = '#,##0'
    dl.txPr = RichText(bodyPr=RichTextProperties(wrap='square'),
                       p=[Paragraph(pPr=ParagraphProperties(defRPr=pr), endParaRPr=pr)])
    return dl


def main():
    pk = openpyxl.load_workbook(PACK)
    if SHEET not in pk.sheetnames:
        print('STOPPED: run patch_v113.py first, there is no %r sheet' % SHEET)
        return
    ws = pk[SHEET]
    if str(ws.cell(27, 1).value).strip() != 'Support Total':
        print('STOPPED at A27: expected %r, found %r. The table has moved.'
              % ('Support Total', ws.cell(27, 1).value))
        return

    # clear anything a previous run left, block and chart alike
    for r in range(TOP, TOP + 10):
        for c in range(1, 14):
            ws.cell(r, c).value = None
    ws._charts = []

    thin = Side(style='thin', color=LINE)
    ws.cell(TOP, 1, 'HOURS BY THE THREE DEPARTMENTS, ROLLING FORECAST').font = Font(
        name='Calibri', bold=True, size=10, color=INK)
    head = TOP + 1
    # every series says hours, so the legend cannot be read as anything else
    for c, h in enumerate(['Department', 'Billable hours',
                           'Non-billable hours', 'Leave hours'], 1):
        cell = ws.cell(head, c, h)
        cell.font = Font(name='Calibri', bold=True, color='FFFFFF', size=10)
        cell.fill = PatternFill('solid', fgColor=INK)
        cell.alignment = Alignment(horizontal='left' if c == 1 else 'right')
        cell.border = Border(bottom=Side(style='medium', color=TEAL))

    out = head + 1
    first = out
    for name, src in ROWS:
        ws.cell(out, 1, name)
        for c, col in enumerate((CH, NC, LV), 2):
            ws.cell(out, c, '=%s%d' % (col, src))
        for c in range(1, 5):
            cell = ws.cell(out, c)
            cell.font = Font(name='Calibri', size=9, color=INK)
            cell.alignment = Alignment(horizontal='left' if c == 1 else 'right')
            if c > 1:
                cell.number_format = '#,##0'
            cell.border = Border(bottom=thin)
        out += 1
    last = out - 1
    ws.cell(out, 1, 'Total')
    for c in range(2, 5):
        ws.cell(out, c, '=SUM(%s%d:%s%d)' % (chr(64 + c), first, chr(64 + c), last))
    for c in range(1, 5):
        cell = ws.cell(out, c)
        cell.font = Font(name='Calibri', bold=True, size=9, color=INK)
        cell.alignment = Alignment(horizontal='left' if c == 1 else 'right')
        if c > 1:
            cell.number_format = '#,##0'
        cell.fill = PatternFill('solid', fgColor=BAND)
        cell.border = Border(bottom=thin)

    note = ws.cell(out + 1, 1, 'All figures are hours.')
    note.font = Font(name='Calibri', italic=True, size=8, color=INK)

    ch = BarChart()
    ch.type = 'bar'                       # horizontal, as the slide has it
    ch.grouping = 'percentStacked'
    ch.overlap = 100
    ch.gapWidth = 45
    ch.title = None                       # the slide heading is the heading
    data = Reference(ws, min_col=2, max_col=4, min_row=head, max_row=last)
    cats = Reference(ws, min_col=1, min_row=first, max_row=last)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    for ser, colour, ink in zip(ch.series, (BILL, NONBILL, LEAVE),
                                ('ffffff', 'ffffff', INK.lower())):
        ser.graphicalProperties = GraphicalProperties(solidFill=colour)
        ser.graphicalProperties.line.noFill = True
        ser.dLbls = labels(ink)
    ch.legend.position = 'b'
    ch.legend.overlay = False
    # for a horizontal bar openpyxl still calls the category axis x_axis
    ch.x_axis.delete = False              # Support, Production, Consulting
    ch.y_axis.delete = True               # the data labels carry the numbers
    ch.x_axis.majorGridlines = None
    ch.y_axis.majorGridlines = None
    ch.x_axis.majorTickMark = 'none'
    ch.x_axis.txPr = RichText(
        bodyPr=RichTextProperties(),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(
            sz=900, latin=DFont(typeface='Calibri'), solidFill=INK.lower())),
            endParaRPr=CharacterProperties(sz=900))])
    # LibreOffice sizes the plot area for about two and a half bars and clips
    # the top one. A manual layout pins it, so all three show.
    ch.layout = Layout(manualLayout=ManualLayout(
        x=0.10, y=0.03, w=0.88, h=0.82, xMode='edge', yMode='edge'))
    ch.width, ch.height = 30, 11          # three bars; the five bar one is the same height
    ws.add_chart(ch, 'A%d' % (out + 3))

    pk.save(PACK)
    print('  block written at %r rows %d-%d' % (SHEET, TOP, out))
    print('  chart anchored at A%d, three bars, percent stacked' % (out + 3))
    print("  snip range '%s'!A%d:D%d for the table" % (SHEET, head, out))


if __name__ == '__main__':
    main()
