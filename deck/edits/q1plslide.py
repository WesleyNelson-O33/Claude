"""Add the quarterly profit and loss by reporting group.

Danica asked for a quarterly P&L per department. The groups go across the top
rather than one slide each, because each group's own block is tall and narrow
and lands on a landscape slide about two inches wide.

Below gross profit the pack splits to these three groups only, not to the five
departments. Events and Video sit inside Production, Integration and
Consulting inside Consulting.

This is a QUARTERLY slide. It sits beside the other quarterly departmental
one, and its commentary stays on the quarter.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from addslide import add

SNIP = '/home/user/Claude/deck/snips/'

BULLETS = [
    'Support is the only group that earns its overhead back, at $73,050 net on a 10.2% margin.',
    'Production holds the best gross margin at 59.2%, then loses it to $325,700 of overhead.',
    'Consulting keeps only 11.0% at the gross line, the weakest of the three.',
    'Production and Consulting are $78,858 down between them for the quarter.',
    'Group EBITDA was $5,774 and the quarter closed as a net loss of $5,808.',
]

NOTES = (
    'A full profit and loss for the quarter, with the three reporting groups across '
    'the top.\n\n'
    'Support earns $718,587, keeps $310,919 at the gross line and still has $73,050 '
    'left after $240,184 of overhead. It is the only group that covers what it is '
    'charged.\n\n'
    'Production looks strongest at the gross line, at 59.2 per cent, but carries '
    '$325,700 of overhead and finishes $52,081 down.\n\n'
    'Consulting is the opposite problem. It bills almost as much as Production, '
    '$445,481, but keeps only 11.0 per cent of it, so $48,782 of gross profit against '
    '$75,653 of overhead leaves it $26,777 down.\n\n'
    'Below gross profit the pack splits to these three groups only. Events and Video '
    'sit inside Production, Integration and Consulting inside Consulting. The five '
    'department split exists down to gross profit, not below it.'
)


def main(src, out, at):
    pres = Presentation(src)
    before = len(pres.slides)
    w, h, strip = add(pres, 5, int(at), 'DEPARTMENTAL PERFORMANCE',
                      'Profit and loss by reporting group, first quarter FY2027',
                      SNIP + 'q1_pl.png', BULLETS, NOTES, 'snip q1_pl')
    print('  q1_pl at index %s, picture %.2f x %.2f, strip %.2f' % (at, w, h, strip))
    pres.save(out)
    print('%d slides in, %d out' % (before, len(pres.slides)))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3])
