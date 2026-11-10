"""Put every slide heading into title case.

Danica asked for each word in a heading to start with a capital.

A word is left exactly as it is when it already carries a capital after the
first letter, or holds a digit, or is not a word at all. That protects FY2027,
FY25, Q1, P&L, EBITDA, DSO and the like, which would otherwise be flattened.

Only the heading is touched. The section eyebrow above it is already upper
case, the commentary is sentence case and stays that way.
"""
import re, sys
from pptx import Presentation

WORD = re.compile(r"[A-Za-z][A-Za-z'&]*")


def fix_word(w):
    if any(ch.isdigit() for ch in w):
        return w
    if any(ch.isupper() for ch in w[1:]):
        return w
    return w[0].upper() + w[1:]


def title_case(text):
    return WORD.sub(lambda m: fix_word(m.group(0)), text)


def main(src, out):
    pres = Presentation(src)
    changed = 0
    for i, s in enumerate(pres.slides, 1):
        for sh in s.shapes:
            if sh.name != 'Text 1' or not sh.has_text_frame:
                continue
            for para in sh.text_frame.paragraphs:
                for run in para.runs:
                    new = title_case(run.text)
                    if new != run.text:
                        print('  %-3d %s' % (i, new))
                        run.text = new
                        changed += 1
    pres.save(out)
    print('\n%d headings changed' % changed)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
