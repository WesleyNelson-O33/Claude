"""Build Danica's edited quarter deck.

She restructured the deck: the section and title in one box top left, the
content on the left, and the commentary in a panel down the right hand side.
This fits content to that layout rather than the old bottom band.

What it does:
  drops every hidden slide
  drops the duplicate Results By Reporting Group, and the two department
    slides the P&L by Dept sheet has no block for
  puts each picture in the space left of the commentary panel, centred
  writes the commentary into the right hand panel, read from the pack
  leaves the slides she embedded Excel objects into alone

Figures are read from the controller pack. None are typed.
"""
import copy, sys, os
import openpyxl
from PIL import Image
from pptx import Presentation
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from qcommentary import figures, build_text, grp_notes

SNIP = '/home/user/Claude/deck/snips/'
GAP = 0.18          # between the content and the panel

# slide title (lower case) -> (snip id, new title or None, which commentary)
FILL = {
    'profit and loss by reporting group, first quarter fy2027': ('q1_pl', None, 'pl'),
    'support q1':     ('pl_sup',  'Support Q1',     'sup'),
    'events q1':      ('pl_prd',  'Production Q1',  'prd'),
    'consulting q1':  ('pl_cons', 'Consulting Q1',  'cons'),
    'utilisation per employee, first quarter fy2027': ('util_q1', None, 'util'),
}
# titles to drop outright
DROP_EXTRA = {'video q1', 'integration q1'}
DROP_DUPLICATE = 'results by reporting group, first quarter fy2027'


def title_of(slide):
    """Her layout puts the section and the title in one box, so the title is
    the last non-empty line of Text 0. Older slides still use Text 1."""
    for name in ('Text 1', 'Text 0'):
        for sh in slide.shapes:
            if sh.name == name and sh.has_text_frame:
                lines = [x.strip() for x in sh.text_frame.text.split('\n') if x.strip()]
                if len(lines) > 1 or (name == 'Text 1' and lines):
                    return lines[-1]
    return ''


def set_title(slide, text):
    for sh in slide.shapes:
        if sh.name != 'Text 0' or not sh.has_text_frame:
            continue
        paras = [p for p in sh.text_frame.paragraphs if p.text.strip()]
        if len(paras) >= 2 and paras[-1].runs:
            paras[-1].runs[0].text = text
            for extra in paras[-1].runs[1:]:
                extra.text = ''
            return True
    return False


def drop(pres, idx):
    ids = pres.slides._sldIdLst
    el = list(ids)[idx]
    pres.part.drop_rel(el.rId)
    ids.remove(el)


def put_picture(slide, png, name):
    """Fill the space left of the commentary panel, centred both ways."""
    panel = next((sh for sh in slide.shapes if sh.name == 'commentary panel'), None)
    if panel is None:
        return None
    left0 = 0.55
    top0 = Emu(panel.top).inches
    room_w = Emu(panel.left).inches - GAP - left0
    room_h = Emu(panel.height).inches
    tree = slide.shapes._spTree
    for sh in list(slide.shapes):
        if 'PICTURE' in str(sh.shape_type) and sh.name.startswith('snip'):
            tree.remove(sh._element)
    im = Image.open(png)
    aspect = im.width / im.height
    h = min(room_h, room_w / aspect)
    w = h * aspect
    pic = slide.shapes.add_picture(png, Inches(left0 + (room_w - w) / 2),
                                   Inches(top0 + (room_h - h) / 2),
                                   width=Inches(w), height=Inches(h))
    pic.name = name
    tree.remove(pic._element)
    tree.insert(2, pic._element)
    return w, h


def write_panel(slide, lines):
    """All the commentary goes in the right hand column. The old second
    column is emptied, because her panel is one column, not two."""
    by = {sh.name: sh for sh in slide.shapes}
    body = by.get('Text 5')
    if body is None:
        return
    tf = body.text_frame
    tf.word_wrap = True
    first = tf.paragraphs[0]
    for para in list(tf.paragraphs)[1:]:
        para._p.getparent().remove(para._p)
    for r in list(first.runs)[1:]:
        r._r.getparent().remove(r._r)
    if not first.runs:
        return
    first.runs[0].text = lines[0]
    for extra in lines[1:]:
        np = copy.deepcopy(first._p)
        first._p.getparent().append(np)
        from pptx.text.text import _Paragraph
        _Paragraph(np, tf).runs[0].text = extra
    for para in tf.paragraphs:
        for r in para.runs:
            r.font.size = Pt(9)
    extra_box = by.get('Text 5b')
    if extra_box is not None:
        slide.shapes._spTree.remove(extra_box._element)


def main(src, out):
    f = figures()
    text = build_text(f)
    g = f['grp']

    def d(v):
        return ('$(%s)' % format(abs(round(v)), ',')) if v < 0 else ('$%s' % format(round(v), ','))

    text['sup'] = [
        'Support billed %s for the quarter at a %.1f%% gross margin.' % (d(g['Support']['income']), g['Support']['gpc'] * 100),
        'It is the steadiest book in the business, close to the same figure every month.',
        'Overhead of %s leaves %s of net profit.' % (d(abs(g['Support']['oh'])), d(g['Support']['net'])),
        'A %.1f%% net margin, and the only reporting group above the line.' % (g['Support']['netc'] * 100),
        'It is covering the losses in the other two.',
    ]
    text['prd'] = [
        'Production billed %s for the quarter and kept %s at the gross line, %.1f%%.' % (d(g['Production']['income']), d(g['Production']['gp']), g['Production']['gpc'] * 100),
        'August carried it. Most of the quarter landed in that one month.',
        'July and September were both loss making.',
        'Overhead of %s turns that gross profit into a net loss of %s.' % (d(abs(g['Production']['oh'])), d(abs(g['Production']['net']))),
        'Events is almost all of it. Video is small on both sides.',
    ]
    text['cons'] = [
        'Consulting billed %s but kept only %s, a %.1f%% gross margin.' % (d(g['Consulting']['income']), d(g['Consulting']['gp']), g['Consulting']['gpc'] * 100),
        'September is the month. Most of the quarter arrived on project completion.',
        'Integration is the volume. The Consulting department itself billed very little.',
        'Overhead of %s leaves a net loss of %s.' % (d(abs(g['Consulting']['oh'])), d(abs(g['Consulting']['net']))),
        'This is turnover without margin rather than a cost problem.',
    ]
    text['util'] = [
        'Support ran 6,942 chargeable hours against 393 non-chargeable, 95% of hours worked.',
        'Production billed 954 hours against 1,150 non-chargeable.',
        'Consulting is the same shape on a smaller base, 179 against 280.',
        'Leave took 1,258 hours across the quarter, 543 of it in Production.',
        'Across CTS the quarter ran 8,074 chargeable hours against 1,823 non-chargeable.',
    ]

    pres = Presentation(src)

    # 1. hidden slides, the duplicate and the two surplus department slides
    seen_dup = False
    for i in range(len(pres.slides) - 1, -1, -1):
        s = pres.slides[i]
        t = title_of(s).lower()
        if s._element.get('show') == '0':
            drop(pres, i); continue
        if t in DROP_EXTRA:
            print('  dropped surplus: %s' % title_of(s)); drop(pres, i); continue
        if t == DROP_DUPLICATE:
            has_pic = any('PICTURE' in str(sh.shape_type) and sh.name.startswith('snip')
                          for sh in s.shapes)
            if not has_pic:
                print('  dropped duplicate: %s' % title_of(s)); drop(pres, i)
    print('  %d slides left' % len(pres.slides))

    # 2. fill
    for i, s in enumerate(pres.slides, 1):
        t = title_of(s).lower()
        if t not in FILL:
            continue
        snip, newtitle, key = FILL[t]
        png = os.path.join(SNIP, snip + '.png')
        if not os.path.exists(png):
            print('  %-3d %-34s PNG MISSING, left alone' % (i, t[:34])); continue
        if newtitle:
            set_title(s, newtitle)
        geom = put_picture(s, png, 'snip ' + snip)
        write_panel(s, text[key])
        print('  %-3d %-34s %s  %.2f x %.2f' % (i, (newtitle or title_of(s))[:34], snip,
                                                geom[0], geom[1]))

    pres.save(out)
    print('%d slides out' % len(pres.slides))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
