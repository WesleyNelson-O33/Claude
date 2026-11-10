"""Round thirteen: the two errors a reader found in the deck.

1. Slide 16 was titled Q2 FY27. It is the first quarter.
2. Nine slides still carried the speaker notes from the September monthly
   review. Three of them, Support, Production and Consulting, carried the
   same Support notes as each other, so two of the three were describing the
   wrong department.

Notes are rebuilt from each slide's own commentary, which is read from the
pack, so they cannot say something the slide does not.
"""
import sys, os, re
from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from herdeck import title_of

STALE = re.compile(r'September 2026|Monthly Business Review|Q1 20\d\d', re.I)
BODY = ('Text 5', 'Text 5b', 'Text 53')


def notes_text(slide):
    if not slide.has_notes_slide:
        return None
    tf = slide.notes_slide.notes_text_frame
    return (tf.text or '') if tf is not None else None


def own_words(slide):
    """A slide's commentary if it has one, otherwise whatever it says on its
    face, so a cover or a section divider still gets notes that match it."""
    by = {sh.name: sh for sh in slide.shapes}
    lines = []
    for nm in BODY:
        b = by.get(nm)
        if b is None or not b.has_text_frame:
            continue
        for p in b.text_frame.paragraphs:
            t = ''.join(r.text for r in p.runs).strip()
            if t:
                lines.append(t)
    if lines:
        return lines
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        if sh.name in ('Slide Number Placeholder 0',):
            continue
        t = sh.text_frame.text.strip().replace('\n', ' ')
        if t and not t.isdigit():
            lines.append(t)
    return lines


def main(src, out):
    pres = Presentation(src)

    # 1. the quarter in the title
    fixed_title = 0
    for i, s in enumerate(pres.slides, 1):
        box = {sh.name: sh for sh in s.shapes}.get('Text 0')
        if box is None or not box.has_text_frame:
            continue
        for p in box.text_frame.paragraphs:
            for r in p.runs:
                if 'Q2 FY27' in r.text:
                    r.text = r.text.replace('Q2 FY27', 'Q1 FY27')
                    fixed_title += 1
                    print('  slide %d title corrected to Q1 FY27' % i)

    # 2. the notes that were left behind from the monthly review
    rewritten = []
    for i, s in enumerate(pres.slides, 1):
        t = notes_text(s)
        if t is None or not STALE.search(t):
            continue
        words = own_words(s)
        if not words:
            print('  slide %d has stale notes and nothing to rebuild them from' % i)
            continue
        s.notes_slide.notes_text_frame.text = '\n'.join(words)
        rewritten.append(i)
    print('  %d titles corrected, notes rewritten on slides %s'
          % (fixed_title, rewritten))

    left = [i for i, s in enumerate(pres.slides, 1)
            if STALE.search(notes_text(s) or '')]
    print('  slides still carrying monthly wording in the notes: %s'
          % (left or 'none'))
    pres.save(out)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
