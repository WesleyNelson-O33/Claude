"""Smaller headings throughout, keeping the hierarchy between them.

The page titles carried most of the weight on a slide and crowded the picture
under them. Each step down is proportional, so a page title is still larger
than a section title's subline and the cover still leads.
"""
import sys
from pptx import Presentation
from pptx.util import Pt

# (layout, shape, old) -> new
RULES = [('CTS_LIGHT', 'Text 1', 27.0, 22.0),   # page titles
         ('CTS_DARK',  'Text 4', 28.0, 23.0),   # section titles
         ('CTS_DARK',  'Text 2', 26.0, 22.0),   # the section number in its tile
         ('CTS_DARK',  'Text 6', 40.0, 30.0),   # the month on the cover
         ('CTS_DARK',  'Text 5', 38.0, 30.0)]   # Thank you

p = Presentation(sys.argv[1])
n = 0
for s in p.slides:
    layout = s.slide_layout.name
    for sh in s.shapes:
        if not sh.has_text_frame:
            continue
        for para in sh.text_frame.paragraphs:
            for r in para.runs:
                if not r.font.size:
                    continue
                for lay, name, old, new in RULES:
                    if layout == lay and sh.name == name and abs(r.font.size.pt - old) < 0.01:
                        r.font.size = Pt(new)
                        n += 1
p.save(sys.argv[2])
print(f'{n} headings reduced')
