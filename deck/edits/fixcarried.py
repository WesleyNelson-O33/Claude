"""Correct the commentary on the two carried over year slides.

Danica flagged the fourth quarter slide. Checked every claim on it against the
data on Xero History rows 179 to 189, and two were wrong.

  "Four of the five fourth quarters before this one lost money or made less
  than 1%."  Three lost money. The other two made 3.3% and 5.4%, so neither
  was under one per cent either.

  "The best was 2020 at 20%, on the lowest revenue of the eleven years."
  2020 billed 943,093. 2016 billed 814,578, so 2020 is the second lowest.

Checking the full year slide beside it turned up one more. "Net margin has sat
between 4% and 5% in four of the last five years" is not what the data says.
Only 2024 at 4.5% falls inside that band. The five run 3.9, 8.0, 4.5, 5.3 and
3.8 per cent.

Everything else on both slides checks out and is left alone.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Pt
from addslide import balance, write, SIZE

FIX = {
    'Revenue and net profit, fourth quarter': [
        'Fourth quarter revenue was $1,685,818, the highest fourth quarter on this chart.',
        'Fourth quarter profit was $65,971 at a 3.9% margin, the first positive one since 2022.',
        'Three of the five fourth quarters before this one lost money.',
        'The worst was 2019, which lost $255,544 on $1,465,028 of revenue.',
        'The best was 2020 at 20.2%, on the second lowest revenue of the eleven years.',
    ],
    'Revenue and net profit, full year': [
        'FY2026 revenue was $6,113,186, down $887,171 on FY2025.',
        'FY2026 net profit was $235,273, against $368,638 the year before.',
        'Net margin has run between 3.8% and 5.3% in four of the last five years.',
        '2023 is the exception at 8.0%. The two strong years are 2018 at 17.0% and 2021 at 14.5%.',
        'Revenue has roughly doubled since 2016 while profit has not.',
    ],
}


def title_of(slide):
    for sh in slide.shapes:
        if sh.name == 'Text 1' and sh.has_text_frame:
            return sh.text_frame.text.strip().split('\n')[0]
    return ''


def main(src, out):
    pres = Presentation(src)
    done = 0
    for i, s in enumerate(pres.slides, 1):
        bullets = FIX.get(title_of(s))
        if not bullets:
            continue
        by = {sh.name: sh for sh in s.shapes}
        cols = list(balance(bullets, 4.05))
        for box_name, items in (('Text 5', cols[0]), ('Text 5b', cols[1])):
            box = by.get(box_name)
            if box is None or not items:
                continue
            write(box, items)
            for para in box.text_frame.paragraphs:
                for r in para.runs:
                    r.font.size = Pt(SIZE)
        print('  %-3d %s' % (i, title_of(s)))
        done += 1
    pres.save(out)
    print('%d slides corrected' % done)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
