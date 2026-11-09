"""Commentary to the bottom, picture as large as the slide allows.

Pages 7, 8 and 19 already read this way. This brings 6 and 9 to 14 into line:
the picture centred under the title and the commentary in a full width strip
below it, in two columns.
"""
import copy, io, sys
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.text.text import _Paragraph

WORK = '/tmp/claude-0/-home-user-Claude/a00ffe59-433c-507b-9006-917d8c60083a/scratchpad/work'
TOP = 1.22            # top of the picture
FLOOR = 5.06          # the strip's bottom edge, clear of the footer
MIN_STRIP = 0.90      # it never gets thinner than this


def fit(bullets, w_in, h_in):
    """Largest size at which the longer column still fits the strip."""
    avail = h_in * 72 * 0.94
    for fs in (9.0, 8.5, 8.0, 7.5, 7.0, 6.5):
        cpl = max(12, int(32 * (10 / fs) * (w_in / 2.20)))
        lines = sum(max(1, -(-len(t) // cpl)) for t in bullets)
        if lines * round(fs * 1.25, 1) + (len(bullets) - 1) * 4 <= avail:
            return fs
    return 6.5

def header_px(im):
    """Height of the navy header band at the top of a snipped table."""
    g = im.convert('L')
    w, h = g.size
    last = 0
    for y in range(min(h, 400)):
        row = [g.getpixel((x, y)) for x in range(0, w, max(1, w // 60))]
        if sum(row) / len(row) < 110:
            last = y
    return last + 1 if last else 0

def drop_rows(im, n_rows, drop):
    """Cut whole table rows out of a snip, for rows that carry no figures."""
    head = header_px(im)
    w, h = im.size
    rh = (h - head) / n_rows
    keep = [i for i in range(n_rows) if i + 1 not in drop]
    out = Image.new('RGB', (w, head + int(round(len(keep) * rh))), 'white')
    out.paste(im.crop((0, 0, w, head)), (0, 0))
    y = head
    for i in keep:
        top = int(round(head + i * rh))
        bot = int(round(head + (i + 1) * rh))
        out.paste(im.crop((0, top, w, bot)), (0, y))
        y += bot - top
    return out

def relayout(slide, tag, crop=None):
    by = {sh.name: sh for sh in slide.shapes}
    pic = next(sh for sh in slide.shapes
               if 'PICTURE' in str(sh.shape_type) and sh.name != 'cts logo')
    im = Image.open(io.BytesIO(pic.image.blob)).convert('RGB')
    before = (pic.width / 914400, pic.height / 914400)
    if crop:
        im = drop_rows(im, *crop)
    path = f'{WORK}/{tag}.png'
    im.save(path)
    w, h = im.size
    aspect = w / h

    tree = slide.shapes._spTree
    tree.remove(pic._element)
    ph = min(FLOOR - MIN_STRIP - 0.12 - TOP, 8.90 / aspect)
    pw = ph * aspect
    new = slide.shapes.add_picture(path, Inches(0.55 + (8.90 - pw) / 2),
                                   Inches(TOP), width=Inches(pw), height=Inches(ph))
    new.name = pic.name
    tree.remove(new._element)
    tree.insert(2, new._element)

    # the strip takes whatever the picture leaves, so nothing sits empty
    PY = min(TOP + ph + 0.12, FLOOR - MIN_STRIP)
    PH = FLOOR - PY
    panel, label = by['commentary panel'], by['Text 4']
    panel.left, panel.top = Inches(0.55), Inches(PY)
    panel.width, panel.height = Inches(8.90), Inches(PH)
    label.left, label.top, label.width = Inches(0.75), Inches(PY + 0.08), Inches(3.0)

    # every bullet on the slide, however many boxes they are spread over
    boxes = [by[n] for n in ('Text 5', 'Text 6') if n in by]
    bullets = [p.text for b in boxes for p in b.text_frame.paragraphs if p.text.strip()]
    for extra in boxes[1:]:
        tree.remove(extra._element)
    body = boxes[0]

    half = 2 if len(bullets) >= 4 else 1
    cols = [bullets[:half], bullets[half:]]
    size = min(fit(c, 4.05, PH - 0.30) for c in cols if c)
    second = copy.deepcopy(body._element)
    tree.append(second)
    body2 = next(sh for sh in slide.shapes if sh._element is second)
    body2.name = 'Text 5b'

    for box, items, x in ((body, cols[0], 0.75), (body2, cols[1], 5.10)):
        box.left, box.top = Inches(x), Inches(PY + 0.27)
        box.width, box.height = Inches(4.05), Inches(PH - 0.30)
        tf = box.text_frame
        first = tf.paragraphs[0]
        for para in list(tf.paragraphs)[1:]:
            para._p.getparent().remove(para._p)
        for r in list(first.runs)[1:]:
            r._r.getparent().remove(r._r)
        first.runs[0].text = items[0]
        for extra in items[1:]:
            np = copy.deepcopy(first._p)
            first._p.getparent().append(np)
            _Paragraph(np, tf).runs[0].text = extra
        for para in tf.paragraphs:
            for r in para.runs:
                r.font.size = Pt(size)
    return before, (pw, ph)

# slide -> (tag, crop). On page 6 the two Other Expenses rows are all dashes.
# On the department pages the seven overhead detail rows are collapsed under
# Total Expenses, which is their total and stays on the table. Shorter means
# the picture can run the full width of the slide, so every remaining line is
# half as large again as it was.
OVERHEAD_DETAIL = {9, 10, 11, 12, 13, 14, 15}
# Page 6 collapses both detail blocks under their own totals: the department
# split under Cost of Sales, which pages 9 to 13 carry in full, and the account
# split under Expenses Overheads, which pages 16 and 17 carry in full.
PL_DETAIL = set(range(2, 10)) | set(range(11, 17)) | {17, 18}
JOBS = {6:  ('pl_result', (27, PL_DETAIL)),
        9:  ('dept_Support', (18, OVERHEAD_DETAIL)),
        10: ('dept_Events', (18, OVERHEAD_DETAIL)),
        11: ('dept_Video', (18, OVERHEAD_DETAIL)),
        12: ('dept_Integration', (18, OVERHEAD_DETAIL)),
        13: ('dept_Consulting', (18, OVERHEAD_DETAIL)),
        14: ('rev_by_dept', None)}

if __name__ == '__main__':
    p = Presentation(sys.argv[1])
    for i, (tag, crop) in JOBS.items():
        before, after = relayout(p.slides[i - 1], tag, crop)
        print(f'slide {i:>3} {tag:16s} picture {before[0]:.2f} x {before[1]:.2f}  ->  '
              f'{after[0]:.2f} x {after[1]:.2f}')
    p.save(sys.argv[2])
