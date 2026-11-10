"""Round six: the commentary panel on the first quarter summary slide.

That slide had a panel and a COMMENTARY heading but no body box at all, so
nothing was ever written into it. This adds the box and fills it.

Every figure is read from the Summary of the controller pack. Rows are found
by their label, not by number, so a row moving does not put a wrong figure on
the slide. The two budgeted jobs are Danica's, from her budget model.
"""
import sys, os, copy
import openpyxl
from pptx import Presentation
from pptx.util import Inches, Emu

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from herdeck import title_of, fill_box
from round5 import to_k

PACK = '/home/user/Claude/pack/CTS Financial Controller Pack - v20.xlsx'
YTD, BUD = 5, 7                     # Summary columns: year to date, budget
WANT = ('Total Income', 'Cost of Sales', 'Expenses Overheds',
        'Gross profit', 'Gross profit %', 'EBITDA', 'Net profit')


def k(v):
    n = round(v / 1000)
    return ('$(%s)k' % format(abs(n), ',')) if n < 0 else ('$%sk' % format(n, ','))


def rows(ws):
    found = {}
    for r in range(1, 220):
        lab = ws.cell(r, 1).value
        if isinstance(lab, str) and lab.strip() in WANT and lab.strip() not in found:
            found[lab.strip()] = r
    missing = [w for w in WANT if w not in found]
    if missing:
        print('STOPPED: no row on the Summary labelled %s' % ', '.join(missing))
        return None
    return found


def donor_body(pres):
    """Borrow a commentary body box from a slide whose panel runs down the
    right, so the new one carries the same type and spacing."""
    for s in pres.slides:
        by = {sh.name: sh for sh in s.shapes}
        panel = by.get('commentary panel')
        if panel is not None and Emu(panel.left).inches > 5.0 and 'Text 5' in by:
            return by['Text 5']
    return None


def main(src, out):
    ws = openpyxl.load_workbook(PACK, data_only=True)['Summary']
    R = rows(ws)
    if R is None:
        return
    g = lambda lab, col: ws.cell(R[lab], col).value
    rev, rev_b = g('Total Income', YTD), g('Total Income', BUD)
    cos, cos_b = abs(g('Cost of Sales', YTD)), abs(g('Cost of Sales', BUD))
    oh, oh_b = abs(g('Expenses Overheds', YTD)), abs(g('Expenses Overheds', BUD))
    gp_pc, gp_pc_b = g('Gross profit %', YTD), g('Gross profit %', BUD)
    eb, eb_b = g('EBITDA', YTD), g('EBITDA', BUD)
    net = g('Net profit', YTD)
    aug = ws.cell(R['Net profit'], 3).value
    jul = ws.cell(R['Net profit'], 2).value
    sep = ws.cell(R['Net profit'], 4).value

    lines = [
        'Revenue was %s against %s budgeted, %s or %.0f%% short.'
        % (k(rev), k(rev_b), k(rev_b - rev), (rev_b - rev) / rev_b * 100),
        'PwC was budgeted at $336k in August and will not now happen.',
        'Bankwest was budgeted at $49k in September and has moved into Q2.',
        'Cost of sales came in %s under budget, not enough to hold the margin.'
        % k(cos_b - cos),
        'Gross margin was %.1f%% against %.1f%% budgeted.' % (gp_pc * 100, gp_pc_b * 100),
        'Overheads were %s, %s under budget.' % (k(oh), k(oh_b - oh)),
        'EBITDA was %s against %s budgeted.' % (k(eb), k(eb_b)),
        'The quarter ends at a net loss of %s.' % k(abs(net)),
        'August was the only month in profit, at %s. July lost %s and September lost %s.'
        % (k(aug), k(abs(jul)), k(abs(sep))),
    ]

    pres = Presentation(src)
    body_src = donor_body(pres)
    done = False
    for i, s in enumerate(pres.slides, 1):
        if not title_of(s).lower().startswith('summary for the first quarter'):
            continue
        by = {sh.name: sh for sh in s.shapes}
        if 'commentary panel' not in by:
            print('  slide %d has no panel' % i)
            return
        box = by.get('Text 5')
        if box is None:
            if body_src is None:
                print('  no donor commentary box found')
                return
            s.shapes._spTree.append(copy.deepcopy(body_src._element))
            box = {sh.name: sh for sh in s.shapes}['Text 5']
            print('  slide %d had no commentary box, one added' % i)
        panel, head = by['commentary panel'], by['Text 4']
        # a cloned shape carries the donor's position, so set all of it, fitted
        # to this slide's own panel rather than to the donor's
        box.left = panel.left + Inches(0.10)
        box.top = head.top + Inches(0.32)
        box.width = panel.width - Inches(0.20)
        box.height = panel.height - (box.top - panel.top) - Inches(0.12)
        fill_box(box, lines)
        for p in box.text_frame.paragraphs:
            for r in p.runs:
                r.text = to_k(r.text)
        s.notes_slide.notes_text_frame.text = '\n'.join(lines)
        print('  slide %d commentary written, %d lines' % (i, len(lines)))
        done = True
        break
    if not done:
        print('  first quarter summary slide not found')
        return
    pres.save(out)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
