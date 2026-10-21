"""A small CTS mark in the top right corner of every light page, in the brand colours.

The deck has no logo image, so this is the same wordmark the cover carries,
"cts" over "Seamless AV", set small and in the light page colours.
"""
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

NAVY, MUTED = RGBColor(0x24, 0x2C, 0x39), RGBColor(0x6B, 0x75, 0x85)
# the three brand colours that carry the dots on the cover pages
TEAL, BLUE, RED = (RGBColor(0x00, 0x53, 0x58), RGBColor(0x6D, 0x71, 0xFF),
                   RGBColor(0xF3, 0x38, 0x44))
MARK = (('c', TEAL), ('t', BLUE), ('s', RED))
X, W = 8.45, 1.00

def add(slide):
    for sh in list(slide.shapes):
        if sh.name in ('cts mark', 'cts mark sub'):
            slide.shapes._spTree.remove(sh._element)

    box = slide.shapes.add_textbox(Inches(X), Inches(0.17), Inches(W), Inches(0.30))
    box.name = 'cts mark'
    tf = box.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.RIGHT
    for letter, colour in MARK:
        run = para.add_run()
        run.text = letter
        run.font.size, run.font.bold, run.font.color.rgb = Pt(15), True, colour

    sub = slide.shapes.add_textbox(Inches(X), Inches(0.42), Inches(W), Inches(0.18))
    sub.name = 'cts mark sub'
    tf = sub.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.RIGHT
    run = para.add_run()
    run.text = 'Seamless AV'
    run.font.size, run.font.color.rgb = Pt(6.5), MUTED

p = Presentation(sys.argv[1])
n = 0
for s in p.slides:
    if s.slide_layout.name != 'CTS_LIGHT':
        continue
    add(s)
    n += 1
p.save(sys.argv[2])
print(f'cts mark added to {n} light pages')
