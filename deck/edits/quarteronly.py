"""Produce the quarter-only copy of the deck.

The deck is one file that serves both reviews. This hides every monthly slide
so what is left reads as the quarterly review on its own. Nothing is deleted,
so unhiding them in PowerPoint brings the monthly review straight back.

Slides are matched on their title, not their position, because every slide
added to the deck shifts the numbering underneath it.

A slide is monthly when it reports September on its own. A slide is kept when
it reports the quarter, the year to date, a position at 30 September, or a run
of full years, because all of those read correctly in a quarterly review.
"""
import sys
from pptx import Presentation

MONTHLY_TITLES = [
    'Performance summary, September 2026',
    'Monthly revenue by department',
    'Revenue against budget and prior years',
    'Monthly revenue, three year history',
    'Revenue by type, average month',
    'Monthly expenditure against budget',
    'Monthly net result, three year comparison',
    'Departmental result in full, September 2026',
    'Gross profit against overhead, September 2026',
    'Results by reporting group, September 2026',
    'Support, September 2026',
    'Events, September 2026',
    'Video, September 2026',
    'Integration, September 2026',
    'Consulting, September 2026',
    'Departmental profit and loss, September 2026',
    'Result by department, above and below the line',
    'Departmental revenue against budget and prior year',
    'Employment expense by function',
    'Chargeable percentage by department, September 2026',
    'Hours mix by department, September 2026',
    'Hours mix by reporting group, September 2026',
]


def title_of(slide):
    for sh in slide.shapes:
        if sh.name == 'Text 1' and sh.has_text_frame:
            return sh.text_frame.text.strip().split('\n')[0]
    return ''


def main(src, out):
    pres = Presentation(src)
    wanted = set(MONTHLY_TITLES)
    seen, hidden = set(), 0
    for s in pres.slides:
        t = title_of(s)
        el = s._element
        if t in wanted:
            el.set('show', '0')
            hidden += 1
            seen.add(t)
        elif el.get('show') is not None:
            del el.attrib['show']
    missing = wanted - seen
    if missing:
        print('WARNING, these titles were not found:')
        for t in sorted(missing):
            print('   ', t)
    total = len(pres.slides)
    pres.save(out)
    print('%d slides, %d hidden as monthly, %d left for the quarter'
          % (total, hidden, total - hidden))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
