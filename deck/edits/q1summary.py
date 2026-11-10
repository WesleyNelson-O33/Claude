"""Add the first quarter summary slide.

Danica asked for a snapshot of the overall summary for the quarter, a cleaner
version of the Summary tab, inserted at page 5 so it opens the financial
performance section.

It is the whole profit and loss on one page: the quarter, its budget, and the
variance in dollars and per cent. The three monthly columns and the rolling
forecast column are hidden, so nothing monthly appears on it.

This is a QUARTERLY slide. Its commentary talks about the quarter only.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from addslide import add

SNIP = '/home/user/Claude/deck/snips/'
AT = 4                 # straight after the section divider, so it becomes page 5

BULLETS = [
    'Revenue for the quarter was $1,624,460 against a budget of $1,996,110, short by $371,650 or 19%.',
    'Gross profit was $632,300, a margin of 38.9% against the 45.7% budgeted.',
    'Overheads were $641,508, which is $48,184 under budget.',
    'EBITDA was $5,774 against a budget of $222,700.',
    'The quarter closed as a net loss of $5,808, against a budgeted profit of $222,700.',
]

NOTES = (
    'The whole quarter on one page, against budget.\n\n'
    'Revenue was $1,624,460 against a budget of $1,996,110. That is $371,650 short, '
    'or 19 per cent.\n\n'
    'Gross profit was $632,300. The margin of 38.9 per cent is nearly seven points '
    'below the 45.7 per cent budgeted, so the shortfall is margin as well as volume.\n\n'
    'Overheads are the one line ahead of plan, at $641,508 against $689,692. That is '
    '$48,184 saved, but it does not cover the gross profit gap.\n\n'
    'EBITDA was $5,774 against a budget of $222,700. After the instant asset write '
    'off of $14,982 the quarter closed as a net loss of $5,808, against a budgeted '
    'profit of $222,700.\n\n'
    'Support is the only reporting group above the line. The departmental slide later '
    'in the deck carries that split.'
)


def main(src, out):
    pres = Presentation(src)
    before = len(pres.slides)
    w, h, strip = add(pres, 5, AT, 'FINANCIAL PERFORMANCE',
                      'Profit and loss summary, first quarter FY2027',
                      SNIP + 'q1_summary.png', BULLETS, NOTES, 'snip q1_summary')
    print('  q1_summary at index %d, picture %.2f x %.2f, strip %.2f' % (AT, w, h, strip))
    pres.save(out)
    print('%d slides in, %d out' % (before, len(pres.slides)))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
