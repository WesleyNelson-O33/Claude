"""Round twelve.

Danica wants the original single quarter top ten slide kept as well as the
three year pair, so it goes back in front of them: this quarter first, then
the history, then the figures.

She also asked for three slides to go: Debtor Days And Collection
Performance, Aged Receivables By Customer, and Ratios Against Industry
Benchmarks.

The commentary on the restored slide is read from the pack's Top 10 Clients
sheet, not retyped.
"""
import sys, os, copy
import openpyxl
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Emu

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from herdeck import title_of, set_title, fill_box

HERE = os.path.dirname(os.path.abspath(__file__))
SNIP = os.path.join(HERE, '..', 'snips')
LOGO = os.path.join(HERE, '..', 'assets', 'cts-logo.png')
PACK = '/home/user/Claude/pack/CTS Financial Controller Pack - v20.xlsx'
RID = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
DROP = ['debtor days and collection',
        'aged receivables by customer',
        'ratios against industry benchmarks']
SOURCE = 'top ten customers, three year comparison'   # the slide to clone from


def k(v):
    return '$%sk' % format(round(v / 1000), ',')


def drop(pres, starts):
    for i, s in enumerate(pres.slides):
        if title_of(s).strip().lower().startswith(starts):
            lst = pres.slides._sldIdLst
            e = list(lst)[i]
            pres.part.drop_rel(e.get(RID))
            lst.remove(e)
            print('  removed slide %d: %s' % (i + 1, title_of(s).strip()))
            return True
    print('  nothing matched %r' % starts)
    return False


def clone_before(pres, src, index):
    new = pres.slides.add_slide(src.slide_layout)
    for sh in list(new.shapes):
        sh._element.getparent().remove(sh._element)
    for sh in src.shapes:
        new.shapes._spTree.append(copy.deepcopy(sh._element))
    lst = pres.slides._sldIdLst
    e = list(lst)[-1]
    lst.remove(e)
    lst.insert(index, e)
    return new


def place(slide, snip):
    """Picture and logo both from file. A cloned slide carries each picture's
    relationship id but not the image behind it."""
    pic = next((x for x in slide.shapes
                if 'PICTURE' in str(x.shape_type) and x.name.startswith('snip ')), None)
    panel = next((x for x in slide.shapes if x.name == 'commentary panel'), None)
    png = os.path.join(SNIP, snip + '.png')
    left = Emu(pic.left).inches
    room_w = 10.0 - 2 * left
    top_lim, bot_lim = 0.94, Emu(panel.top).inches - 0.12
    im = Image.open(png)
    w, h = room_w, room_w * im.height / im.width
    if h > bot_lim - top_lim:
        h = bot_lim - top_lim
        w = h * im.width / im.height
    tree = slide.shapes._spTree
    tree.remove(pic._element)
    new = tree and slide.shapes.add_picture(
        png, Inches(left + (room_w - w) / 2),
        Inches(top_lim + (bot_lim - top_lim - h) / 2),
        width=Inches(w), height=Inches(h))
    new.name = 'snip ' + snip
    tree.remove(new._element)
    tree.insert(2, new._element)
    logo = next((x for x in slide.shapes if x.name == 'cts logo'), None)
    if logo is not None:
        l, t, lw, lh = logo.left, logo.top, logo.width, logo.height
        tree.remove(logo._element)
        slide.shapes.add_picture(LOGO, l, t, width=lw, height=lh).name = 'cts logo'
    return w, h


def main(src, out):
    ws = openpyxl.load_workbook(PACK, data_only=True)['Top 10 Clients']
    rows = {ws.cell(r, 7).value: r for r in range(7, 17)}     # chart label -> row
    ytd = lambda r: ws.cell(r, 4).value
    top3 = sum(ytd(r) for r in range(7, 10))
    top10, others, total = ytd(17), ytd(18), ytd(19)
    deloitte = ytd(rows['Deloitte Services'])
    aware_row = rows['Aware Super']
    aware_ytd, aware_sep = ytd(aware_row), ws.cell(aware_row, 3).value

    pres = Presentation(src)
    # The new slide is added before the three are dropped. Doing it the other
    # way round frees slide parts and the part added next reuses one of the
    # freed names, which leaves a file python-pptx still reads but PowerPoint
    # and LibreOffice refuse to open.
    src_slide, idx = None, None
    for i, s in enumerate(pres.slides):
        if title_of(s).strip().lower().startswith(SOURCE):
            src_slide, idx = s, i
            break
    if src_slide is None:
        print('  the three year slide was not found, nothing added')
        pres.save(out)
        return

    new = clone_before(pres, src_slide, idx)
    set_title(new, 'Top Ten Customers By Revenue, Q1 FY27')
    got = place(new, 'top10')
    by = {sh.name: sh for sh in new.shapes}
    fill_box(by['Text 5'], [
        'The top three customers are %.0f%% of the quarter, %s of %s.'
        % (top3 / total * 100, k(top3), k(total)),
        'The top ten are %.0f%%, leaving %.0f%% across everyone else.'
        % (top10 / total * 100, others / total * 100),
        'Deloitte alone is %.0f%% at %s. That is the concentration risk.'
        % (deloitte / total * 100, k(deloitte)),
    ])
    fill_box(by['Text 5b'], [
        'Aware Super is %s, and %s of that landed in September.'
        % (k(aware_ytd), k(aware_sep)),
        'Revenue falls away quickly after the largest few, so the wider base matters.',
    ])
    try:
        new.notes_slide.notes_text_frame.text = '\n'.join(
            ''.join(r.text for r in p.runs)
            for nm in ('Text 5', 'Text 5b')
            for p in by[nm].text_frame.paragraphs)
    except AttributeError:
        pass
    print('  slide %d restored: Top Ten Customers By Revenue, Q1 FY27 '
          '(%.2f x %.2f)' % (idx + 1, got[0], got[1]))
    for d in DROP:
        drop(pres, d)
    pres.save(out)
    print('  %d slides out' % len(pres.slides))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
