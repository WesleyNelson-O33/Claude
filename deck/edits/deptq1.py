"""Add the quarterly departmental slide.

The departmental section was entirely September. Danica asked for the
departments on a quarter basis as well, so this adds the quarter version of
"Results by reporting group" and sits it directly after the monthly one, which
is how the deck pairs a monthly slide with its quarterly twin.

The figures come from the year to date departmental table on 'P&L by Dept',
which at September is the first quarter. Income, gross profit, EBITDA and net
profit all tie to the Summary tab.

This is a QUARTERLY slide. Its commentary talks about the quarter only and
never refers across to the month, because the monthly slide beside it may be
hidden when the quarter is presented.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from addslide import add

SNIP = '/home/user/Claude/deck/snips/'

EYEBROW = 'DEPARTMENTAL PERFORMANCE'
TITLE   = 'Results by reporting group, first quarter FY2027'
AT      = 23          # straight after "Results by reporting group, September 2026"

BULLETS = [
    'Revenue for the quarter was $1,624,460 and gross profit $632,300, a margin of 38.9%.',
    'Support is the only reporting group above the line, earning $73,050 on a 10.2% net margin.',
    'Production lost $52,081 across the three months and Consulting lost $26,777.',
    'Group EBITDA was $5,774 and the quarter closed as a net loss of $5,808.',
    'Support carries the quarter. The other two groups are $78,858 down between them.',
]

NOTES = (
    'This is the quarter on the same cut as the monthly reporting group slide.\n\n'
    'Revenue for the three months was $1,624,460 and gross profit $632,300, which is '
    'a margin of 38.9 per cent. At group level the quarter made $5,774 of EBITDA and '
    'closed as a net loss of $5,808.\n\n'
    'Support is the only reporting group above the line. It earned $73,050 on revenue '
    'of $718,587, a net margin of 10.2 per cent. Production lost $52,081 on revenue of '
    '$460,391 and Consulting lost $26,777 on revenue of $445,481. Between them those '
    'two are $78,858 down, which is what Support is covering.\n\n'
    'Consulting is the one to watch on volume. It billed almost as much as Production '
    'for the quarter but kept only $48,782 of gross profit against Production\'s '
    '$272,628.\n\n'
    'Figures come from the year to date departmental table in the controller pack and '
    'tie to the Summary tab.'
)


def main(src, out):
    pres = Presentation(src)
    before = len(pres.slides)
    w, h, strip = add(pres, 5, AT, EYEBROW, TITLE, SNIP + 'dept_q1.png',
                      BULLETS, NOTES, 'snip dept_q1')
    print('  dept_q1 at index %d, picture %.2f x %.2f, strip %.2f' % (AT, w, h, strip))
    pres.save(out)
    print('%d slides in, %d out' % (before, len(pres.slides)))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
