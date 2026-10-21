"""The same three dots as the cover, moved to a different position on each
cover and section page. Same sizes, same overlap, different place."""
import copy, sys
from pptx import Presentation
from pptx.util import Inches
from lxml import etree

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
TEAL, BLUE, RED = '3FD0C9', '6D71FF', 'F33844'

# the cover's trio, as offsets from the group's own corner
TRIO = [(TEAL, 0.00, 0.00, 2.35), (BLUE, 1.05, 1.00, 1.95), (RED, 0.25, 1.95, 2.45)]

# slide -> where that group sits
PLACE = {4:  (6.90, 0.20),
         15: (7.45, 1.05),
         20: (6.70, 1.00),
         24: (7.60, 0.30),
         28: (6.60, 1.30),
         31: (7.20, 0.65),
         32: (5.30, 0.40)}
ALPHA = {32: 82000}

def is_dot(shape):
    n = shape.name.lower()
    return n.startswith(('divider circle', 'closing circle', 'dot ')) or n == 'circle'

def main(src, out):
    p = Presentation(src)
    for i, s in enumerate(p.slides, 1):
        if i not in PLACE:
            continue
        tree = s.shapes._spTree
        template = None
        for sh in list(s.shapes):
            if is_dot(sh):
                if template is None:
                    template = copy.deepcopy(sh._element)
                tree.remove(sh._element)
        if template is None:
            continue
        gx, gy = PLACE[i]
        alpha = ALPHA.get(i, 55000)
        for j, (colour, dx, dy, d) in enumerate(TRIO):
            el = copy.deepcopy(template)
            tree.append(el)
            shape = [sh for sh in s.shapes if sh._element is el][0]
            shape.name = f'dot {j}'
            shape.left, shape.top = Inches(gx + dx), Inches(gy + dy)
            shape.width = shape.height = Inches(d)
            clr = shape._element.spPr.find(f'{{{A}}}solidFill/{{{A}}}srgbClr')
            clr.set('val', colour)
            al = clr.find(f'{{{A}}}alpha')
            if al is None:
                al = etree.SubElement(clr, f'{{{A}}}alpha')
            al.set('val', str(alpha))
        for sh in list(s.shapes):
            if sh.name.startswith('dot '):
                tree.remove(sh._element)
                tree.insert(2, sh._element)
    p.save(out)
    print(f'the cover trio placed on {len(PLACE)} pages, a different position on each')

main(sys.argv[1], sys.argv[2])
