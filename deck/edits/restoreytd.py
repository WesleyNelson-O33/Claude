"""Put the year to date slide back the way it was, and give the departments
their own slide again.

Danica asked for the two to be one slide, so they were merged. Seeing it, she
asked for the year to date slide back as it was. It is restored: FY25, FY26,
FY27 and the budget, out to June, so the budget trajectory for the rest of the
year is still on the page.

The department split goes back to its own slide directly after it, which is
where it was before the merge.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Pt
from pptx import Presentation as _P
from addslide import balance, write, SIZE, add

SNIP = '/home/user/Claude/deck/snips/'
MERGED = 'year to date revenue by department and against budget'
RESTORED = 'Year To Date Revenue Against Budget'

YEAR_LINES = [
    'Year to date revenue is $1,624k, ahead of both completed years at this point in the year.',
    'It is $372k or 19% behind budget. Almost all of that opened in August, and September closed $148k of it back.',
    'FY26 lost its year in December and January, $1,045k behind FY25 across those two months.',
    'The full year budget of $6,305k is itself 10% below what FY25 actually billed.',
]

DEPT_LINES = [
    'Support billed $718,587 year to date, the largest of the five departments.',
    'Integration finished the quarter at $441,414 and Events at $420,121.',
    'Integration started the year behind Events and overtook it on project completion.',
    'Video billed $40,270 year to date and Consulting $4,068.',
    'The five together make the $1,624,460 of year to date revenue.',
]

DEPT_NOTES = (
    'The same year to date cut as the slide before, split by department.\n\n'
    'Support is the largest of the five at $718,587 and the steadiest. It adds roughly '
    'the same amount each month, which is what a contract book should look like.\n\n'
    'Integration and Events finish the quarter close together, at $441,414 and '
    '$420,121. They get there differently. Events builds early and flattens, '
    'Integration is slow for two months then jumps on project completion.\n\n'
    'Video at $40,270 and Consulting at $4,068 are both small enough that they do not '
    'move the year to date line.\n\n'
    'The five add to $1,624,460, which is the year to date revenue on the previous '
    'slide.'
)


def title_of(slide):
    for sh in slide.shapes:
        if sh.name == 'Text 1' and sh.has_text_frame:
            return sh.text_frame.text.strip().split('\n')[0]
    return ''


def main(src, out):
    pres = Presentation(src)
    at = None
    for i, s in enumerate(pres.slides):
        if title_of(s).lower() != MERGED:
            continue
        by = {sh.name: sh for sh in s.shapes}
        runs = by['Text 1'].text_frame.paragraphs[0].runs
        runs[0].text = RESTORED
        for extra in runs[1:]:
            extra.text = ''
        cols = list(balance(YEAR_LINES, 4.05))
        for name, items in (('Text 5', cols[0]), ('Text 5b', cols[1])):
            box = by.get(name)
            if box is None or not items:
                continue
            write(box, items)
            for para in box.text_frame.paragraphs:
                for r in para.runs:
                    r.font.size = Pt(SIZE)
        at = i + 1
        print('  restored page %d to %s' % (i + 1, RESTORED))
        break

    if at is None:
        print('  merged slide not found, nothing restored')
        pres.save(out)
        return

    # the template is found rather than numbered, because every slide inserted
    # into the deck moves the index of the one that used to be page 6
    tpl = next(i for i, sl in enumerate(pres.slides)
               if {'commentary panel', 'Text 4', 'Text 5'} <= {sh.name for sh in sl.shapes}
               and any('PICTURE' in str(sh.shape_type) and sh.name.startswith('snip ')
                       for sh in sl.shapes))
    print('  template is page %d' % (tpl + 1))
    w, h, strip = add(pres, tpl, at, 'FINANCIAL PERFORMANCE',
                      'Year To Date Revenue By Department',
                      SNIP + 'ytd_dept.png', DEPT_LINES, DEPT_NOTES, 'snip ytd_dept')
    print('  department slide added at page %d, picture %.2f x %.2f' % (at + 1, w, h))
    pres.save(out)
    print('%d slides out' % len(pres.slides))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
