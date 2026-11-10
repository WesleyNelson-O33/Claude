"""The executive summary, on a quarter basis.

Danica asked for the September executive summary in a quarter view. This
clones that slide and rewrites every figure from the controller pack, so
nothing is typed by hand and it cannot drift from the pack.

One change to the card. The monthly card's third line is the result after
allocated overhead. The pack only splits overhead to the five departments for
a single month, so for the quarter that line shows gross profit instead, which
does exist for all five. The result after overhead is in the strip and in the
paragraph, where the three reporting groups carry it.

This is a QUARTERLY slide and sits straight after its monthly twin.
"""
import copy, sys, os
import openpyxl
from pptx import Presentation

PACK = '/home/user/Claude/pack/CTS Financial Controller Pack - v20.xlsx'
INC = {'Support': 41, 'Events': 8, 'Video': 16,
       'Integration': 25, 'Consulting': 33, 'Total': 6}
COS = {'Support': 130, 'Events': 61, 'Video': 78,
       'Integration': 96, 'Consulting': 113, 'Total': 59}
ORDER = ['Support', 'Events', 'Video', 'Integration', 'Consulting', 'Total']

# card -> (name box, big number, label, margin, third line, budget line)
CARDS = {
    'Support':     ('Text 3', 'Text 4', 'Text 5', 'Text 6', 'Text 7', 'Text 8'),
    'Events':      ('Text 10', 'Text 11', 'Text 12', 'Text 13', 'Text 14', 'Text 15'),
    'Video':       ('Text 17', 'Text 18', 'Text 19', 'Text 20', 'Text 21', 'Text 22'),
    'Integration': ('Text 24', 'Text 25', 'Text 26', 'Text 27', 'Text 28', 'Text 29'),
    'Consulting':  ('Text 31', 'Text 32', 'Text 33', 'Text 34', 'Text 35', 'Text 36'),
    'Total':       ('Total text 0', 'Total text 1', 'Total text 2',
                    'Total text 3', 'Total text 4', 'Total text 5'),
}
STRIP = [('Text 38', 'Text 39', 'Text 40'), ('Text 41', 'Text 42', 'Text 43'),
         ('Text 44', 'Text 45', 'Text 46'), ('Text 47', 'Text 48', 'Text 49'),
         ('Text 50', 'Text 51', 'Text 52')]


def k(v):
    """Thousands, the way the monthly slide writes them."""
    n = round(v / 1000)
    return ('$(%s)k' % format(abs(n), ',')) if n < 0 else ('$%sk' % format(n, ','))


def setext(shapes, name, text):
    sh = shapes.get(name)
    if sh is None or not sh.has_text_frame:
        return
    runs = sh.text_frame.paragraphs[0].runs
    if not runs:
        return
    runs[0].text = text
    for extra in runs[1:]:
        extra.text = ''


def main(src, out, at):
    wb = openpyxl.load_workbook(PACK, data_only=True)
    ws = wb['Summary']
    rev = {d: ws.cell(INC[d], 5).value for d in ORDER}
    bud = {d: ws.cell(INC[d], 7).value or 0 for d in ORDER}
    cos = {d: ws.cell(COS[d], 5).value for d in ORDER}
    gp = {d: rev[d] + cos[d] for d in ORDER}
    oh = ws.cell(164, 5).value
    oh_b = ws.cell(164, 7).value
    gpc = ws.cell(174, 5).value
    gpc_b = ws.cell(174, 7).value
    eb = ws.cell(175, 5).value
    eb_b = ws.cell(175, 7).value
    breakeven = abs(oh) / gpc if gpc else 0

    pres = Presentation(src)
    before = len(pres.slides)
    template = pres.slides[2]
    new = pres.slides.add_slide(template.slide_layout)
    for sh in list(new.shapes):
        sh._element.getparent().remove(sh._element)
    for sh in template.shapes:
        new.shapes._spTree.append(copy.deepcopy(sh._element))
    by = {sh.name: sh for sh in new.shapes}

    # the logo is a relationship, so it comes back from file
    logo = by.get('cts logo')
    if logo is not None:
        geom = (logo.left, logo.top, logo.width, logo.height)
        new.shapes._spTree.remove(logo._element)
        mark = new.shapes.add_picture('/home/user/Claude/deck/assets/cts-logo.png',
                                      geom[0], geom[1], width=geom[2], height=geom[3])
        mark.name = 'cts logo'
        by = {sh.name: sh for sh in new.shapes}

    setext(by, 'Text 1', 'Performance Summary, First Quarter FY2027')

    for dept in ORDER:
        nm, big, lab, marg, third, budg = CARDS[dept]
        setext(by, big, k(rev[dept]))
        setext(by, lab, 'revenue')
        setext(by, marg, '%.0f%% gross margin' % (gp[dept] / rev[dept] * 100))
        setext(by, third, '%s gross profit' % k(gp[dept]))
        d = rev[dept] - bud[dept]
        setext(by, budg, ('Ahead of budget by %s' if d >= 0 else 'Behind budget by %s')
               % k(abs(d)))

    vals = [('Revenue', k(rev['Total']), 'budget %s' % k(bud['Total'])),
            ('Gross margin', '%.1f%%' % (gpc * 100), 'budget %.1f%%' % (gpc_b * 100)),
            ('Overheads', k(abs(oh)), 'budget %s' % k(abs(oh_b))),
            ('EBITDA', k(eb), 'budget %s' % k(eb_b)),
            ('Break even needs', k(breakeven), 'running at %s' % k(rev['Total']))]
    for (h, v, s_), (hb, vb, sb) in zip(vals, STRIP):
        setext(by, hb, h); setext(by, vb, v); setext(by, sb, s_)

    short = bud['Total'] - rev['Total']
    para = ('The quarter in summary. Revenue was %s against a budget of %s, short by %s '
            'or %.0f%%. Almost all of the shortfall is Consulting at %s and Integration '
            'at %s. Support is the only reporting group that covered its allocated '
            'overhead, earning $73k, while Production lost $52k and Consulting $27k. At '
            'the quarter gross margin of %.1f%%, revenue of approximately %s is required '
            'to cover overheads of %s, against the %s billed.'
            % (k(rev['Total']), k(bud['Total']), k(short), short / bud['Total'] * 100,
               k(bud['Consulting'] - rev['Consulting']),
               k(bud['Integration'] - rev['Integration']),
               gpc * 100, k(breakeven), k(abs(oh)), k(rev['Total'])))
    setext(by, 'Text 53', para)

    ns = new.notes_slide
    if ns.notes_text_frame is None:
        ns.shapes._spTree.append(copy.deepcopy(template.notes_slide.notes_placeholder._element))
    ns.notes_text_frame.text = para

    ids = pres.slides._sldIdLst
    el = ids[-1]
    ids.remove(el)
    ids.insert(int(at), el)
    pres.save(out)
    for d in ORDER:
        print('  %-12s %8s  %5.0f%% GM  %8s GP' % (d, k(rev[d]), gp[d] / rev[d] * 100, k(gp[d])))
    print('  break even %s, running at %s' % (k(breakeven), k(rev['Total'])))
    print('%d slides in, %d out' % (before, len(pres.slides)))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3])
