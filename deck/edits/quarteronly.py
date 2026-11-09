"""Produce the quarter-only copy of the deck.

The deck is one file that serves both reviews. This hides every monthly slide
so what is left reads as the quarterly review on its own. Nothing is deleted,
so unhiding them in PowerPoint brings the monthly review straight back.

A slide is monthly when it reports September on its own. A slide is kept when
it reports the quarter, the year to date, a position at 30 September, or a
run of full years, because all of those read correctly in a quarterly review.
"""
import sys
from pptx import Presentation

# 1-based slide numbers in the 52 slide deck.
MONTHLY = {
    3,                               # Performance summary, September 2026
    6, 7, 9, 11, 13, 15,             # the monthly financial performance cuts
    21, 22, 23,                      # departmental, September
    25, 26, 27, 28, 29,              # the five department pages, September
    30, 31, 32,                      # departmental P&L and budget, September
    37,                              # employment expense by function, September
    43, 44, 45,                      # utilisation, September
}


def main(src, out):
    pres = Presentation(src)
    total = len(pres.slides)
    hidden = 0
    for i, s in enumerate(pres.slides, 1):
        el = s._element
        if i in MONTHLY:
            el.set('show', '0')
            hidden += 1
        elif el.get('show') is not None:
            del el.attrib['show']
    pres.save(out)
    print('%d slides, %d hidden as monthly, %d left for the quarter'
          % (total, hidden, total - hidden))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
