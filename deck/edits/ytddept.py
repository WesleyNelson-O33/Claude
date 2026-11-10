"""Add the year to date revenue by department slide.

Danica asked for department lines on the year to date revenue graph. They
cannot go on that graph. Its year lines run to seven million and the
departments to well under one, so on a single picture the departments flatten
onto the axis and the legend collides with the FY27 line colour.

So they get the same cut on their own chart, and the slide sits directly after
the year one.

This is a YEAR TO DATE slide, which at September is the quarter, so it belongs
with the quarterly set. Its commentary stays on the year to date.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from addslide import add

SNIP = '/home/user/Claude/deck/snips/'
AT = 10                 # straight after "Year to date revenue against budget"

BULLETS = [
    'Support billed $718,587 year to date, the largest of the five departments.',
    'Integration finished the quarter at $441,414 and Events at $420,121.',
    'Integration started the year behind Events and overtook it on project completion.',
    'Video billed $40,270 year to date and Consulting $4,068.',
    'The five together make the $1,624,460 of year to date revenue.',
]

NOTES = (
    'The same year to date cut as the slide before, split by department.\n\n'
    'Support is the largest of the five at $718,587 and the steadiest. It adds '
    'roughly the same amount each month, which is what a contract book should look '
    'like.\n\n'
    'Integration and Events finish the quarter close together, at $441,414 and '
    '$420,121. They get there differently. Events builds early and flattens, '
    'Integration is slow for two months then jumps on project completion.\n\n'
    'Video at $40,270 and Consulting at $4,068 are both small enough that they do '
    'not move the year to date line.\n\n'
    'The five add to $1,624,460, which is the year to date revenue on the previous '
    'slide.'
)


def main(src, out):
    pres = Presentation(src)
    before = len(pres.slides)
    w, h, strip = add(pres, 5, AT, 'FINANCIAL PERFORMANCE',
                      'Year to date revenue by department',
                      SNIP + 'ytd_dept.png', BULLETS, NOTES, 'snip ytd_dept')
    print('  ytd_dept at index %d, picture %.2f x %.2f, strip %.2f' % (AT, w, h, strip))
    pres.save(out)
    print('%d slides in, %d out' % (before, len(pres.slides)))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
