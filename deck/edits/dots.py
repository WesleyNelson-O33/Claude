"""All three brand colours on every cover and section page, arranged
differently on each one."""
import copy, sys
from pptx import Presentation
from pptx.util import Inches
from lxml import etree

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
TEAL, BLUE, RED = '3FD0C9', '6D71FF', 'F33844'

# slide -> list of (colour, x, y, diameter). Drawn in order, so the last sits
# on top. Each page gets its own composition.
PLAN = {
    4:  [(TEAL, 6.60, 2.45, 2.35), (BLUE, 8.05, 1.35, 1.70), (RED, 7.35, 3.55, 2.05)],
    15: [(BLUE, 7.25, 0.30, 1.60), (TEAL, 6.70, 1.65, 2.45), (RED, 8.25, 3.25, 2.25)],
    20: [(RED, 6.85, 2.15, 2.95), (TEAL, 8.35, 0.75, 1.35), (BLUE, 6.95, 0.90, 1.05)],
    24: [(TEAL, 6.15, 3.35, 1.75), (BLUE, 7.55, 3.00, 2.15), (RED, 8.95, 3.65, 1.65)],
    28: [(BLUE, 6.95, 1.15, 1.95), (RED, 8.55, 1.95, 2.35), (TEAL, 7.25, 3.15, 1.85)],
    31: [(TEAL, 7.55, 0.55, 1.55), (BLUE, 8.35, 1.85, 2.05), (RED, 7.85, 3.35, 2.45)],
    # the closing page keeps three, re-set so it does not repeat the cover
    32: [(RED, 5.35, 0.45, 2.10), (TEAL, 6.60, 1.55, 2.60), (BLUE, 5.95, 2.95, 2.20)],
}
ALPHA = {32: 82000}          # the closing page matches the cover's weight

def is_dot(shape):
    name = shape.name.lower()
    return name.startswith(('divider circle', 'closing circle')) or name == 'circle'

def main(src, out):
    p = Presentation(src)
    for i, s in enumerate(p.slides, 1):
        if i not in PLAN:
            continue
        tree = s.shapes._spTree
        template = None
        for sh in list(s.shapes):
            if is_dot(sh):
                template = template or copy.deepcopy(sh._element)
                tree.remove(sh._element)
        if template is None:
            continue
        alpha = ALPHA.get(i, 55000)
        for j, (colour, x, y, d) in enumerate(PLAN[i]):
            el = copy.deepcopy(template)
            tree.append(el)
            shape = [sh for sh in s.shapes if sh._element is el][0]
            shape.name = f'dot {j}'
            shape.left, shape.top = Inches(x), Inches(y)
            shape.width = shape.height = Inches(d)
            clr = shape._element.spPr.find(f'{{{A}}}solidFill/{{{A}}}srgbClr')
            clr.set('val', colour)
            al = clr.find(f'{{{A}}}alpha')
            if al is None:
                al = etree.SubElement(clr, f'{{{A}}}alpha')
            al.set('val', str(alpha))
        # the dots belong behind the wording
        for sh in list(s.shapes):
            if sh.name.startswith('dot '):
                tree.remove(sh._element)
                tree.insert(2, sh._element)
    p.save(out)
    print(f'three dots placed on {len(PLAN)} cover and section pages')

main(sys.argv[1], sys.argv[2])
