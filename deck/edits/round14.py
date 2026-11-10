"""Round fourteen: the profit per employee slide.

Its chart now reads FY22 to FY26, so the commentary should too.

The slide also carried its heading twice, once in the title box and again in
a second box straight underneath. The slide heading is the heading.
"""
import sys, os, re
from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from herdeck import title_of

TARGET = 'profit per employee'
YEARS = {'2023': 'FY23', '2025': 'FY25', '2022': 'FY22',
         '2024': 'FY24', '2026': 'FY26'}
BARE = re.compile(r'(?<![$\w])(202[2-6])(?!\d)')


def main(src, out):
    pres = Presentation(src)
    for i, s in enumerate(pres.slides, 1):
        if not title_of(s).strip().lower().startswith(TARGET):
            continue
        by = {sh.name: sh for sh in s.shapes}

        # the heading was in the title box and again underneath it
        t0, t1 = by.get('Text 0'), by.get('Text 1')
        if t0 is not None and t1 is not None and t1.has_text_frame:
            if t1.text_frame.text.strip().lower() == \
                    t0.text_frame.text.split('|')[-1].strip().lower():
                t1._element.getparent().remove(t1._element)
                print('  slide %d: the repeated heading removed' % i)

        n = 0
        for nm in ('Text 5', 'Text 5b'):
            b = by.get(nm)
            if b is None:
                continue
            for p in b.text_frame.paragraphs:
                for r in p.runs:
                    new = BARE.sub(lambda m: YEARS[m.group(1)], r.text)
                    if new != r.text:
                        r.text = new
                        n += 1
        print('  slide %d: %d runs now read the financial year' % (i, n))
        try:
            s.notes_slide.notes_text_frame.text = '\n'.join(
                ''.join(r.text for r in p.runs)
                for nm in ('Text 5', 'Text 5b') if by.get(nm) is not None
                for p in by[nm].text_frame.paragraphs)
        except AttributeError:
            pass
        break
    pres.save(out)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
