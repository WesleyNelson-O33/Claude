"""Turn the executive summary from the month into the quarter, in place.

Danica's quarter deck still carried the September executive summary. Same
slide, same layout, quarter figures.

Everything is read from the controller pack. The per department revenue,
budget, gross profit and result after overhead come from the quarterly
department block the pack's own allocation writes at Overhead Split row 300.
The company strip comes from the Summary.
"""
import sys, os
import openpyxl
from pptx import Presentation

PACK = '/home/user/Claude/pack/CTS Financial Controller Pack - v20.xlsx'
TOP = 300
DEPTS = ['Support', 'Events', 'Video', 'Integration', 'Consulting', 'Total']
CARDS = {
    'Support':     ('Text 4', 'Text 6', 'Text 7', 'Text 8'),
    'Events':      ('Text 11', 'Text 13', 'Text 14', 'Text 15'),
    'Video':       ('Text 18', 'Text 20', 'Text 21', 'Text 22'),
    'Integration': ('Text 25', 'Text 27', 'Text 28', 'Text 29'),
    'Consulting':  ('Text 32', 'Text 34', 'Text 35', 'Text 36'),
    'Total':       ('Total text 1', 'Total text 3', 'Total text 4', 'Total text 5'),
}
STRIP = [('Text 39', 'Text 40'), ('Text 42', 'Text 43'), ('Text 45', 'Text 46'),
         ('Text 48', 'Text 49'), ('Text 51', 'Text 52')]


def k(v):
    n = round(v / 1000)
    return ('$(%s)k' % format(abs(n), ',')) if n < 0 else ('$%sk' % format(n, ','))


def setext(by, name, text):
    sh = by.get(name)
    if sh is None or not sh.has_text_frame:
        return
    runs = sh.text_frame.paragraphs[0].runs
    if not runs:
        return
    runs[0].text = text
    for extra in runs[1:]:
        extra.text = ''


def main(src, out):
    wb = openpyxl.load_workbook(PACK, data_only=True)
    ws, sm = wb['Overhead Split'], wb['Summary']
    d = {}
    for i, name in enumerate(DEPTS):
        r = TOP + 2 + i
        got = ws.cell(r, 1).value
        if str(got).strip() != name:
            print('STOPPED: expected %s at row %d, found %r' % (name, r, got))
            return
        d[name] = {'rev': ws.cell(r, 2).value, 'bud': ws.cell(r, 3).value,
                   'gp': ws.cell(r, 4).value, 'res': ws.cell(r, 6).value}
    oh, oh_b = abs(sm.cell(164, 5).value), abs(sm.cell(164, 7).value)
    gpc, gpc_b = sm.cell(174, 5).value, sm.cell(174, 7).value
    eb, eb_b = sm.cell(175, 5).value, sm.cell(175, 7).value
    breakeven = oh / gpc if gpc else 0

    pres = Presentation(src)
    done = False
    for i, s in enumerate(pres.slides, 1):
        by = {sh.name: sh for sh in s.shapes}
        t1 = by.get('Text 1')
        if t1 is None or 'Performance Summary' not in t1.text_frame.text:
            continue
        setext(by, 'Text 1', 'Performance Summary, First Quarter FY2027')
        for name, (big, marg, third, budg) in CARDS.items():
            e = d[name]
            setext(by, big, k(e['rev']))
            setext(by, marg, '%.0f%% gross margin' % (e['gp'] / e['rev'] * 100))
            res = e['res']
            setext(by, third, '%s%s after overhead' % ('+' if res >= 0 else '', k(res)))
            diff = e['rev'] - e['bud']
            setext(by, budg, ('Ahead of budget by %s' if diff >= 0 else 'Behind budget by %s')
                   % k(abs(diff)))
        vals = [(k(d['Total']['rev']), 'budget %s' % k(d['Total']['bud'])),
                ('%.1f%%' % (gpc * 100), 'budget %.1f%%' % (gpc_b * 100)),
                (k(oh), 'budget %s' % k(oh_b)),
                (k(eb), 'budget %s' % k(eb_b)),
                (k(breakeven), 'running at %s' % k(d['Total']['rev']))]
        for (v, sub), (vb, sb) in zip(vals, STRIP):
            setext(by, vb, v)
            setext(by, sb, sub)
        short = d['Total']['bud'] - d['Total']['rev']
        para = ('The quarter in summary. Revenue was %s against a budget of %s, short by %s '
                'or %.0f%%. Almost all of the shortfall is Integration at %s and Consulting '
                'at %s. Support is the only department covering its allocated overhead, at '
                '%s, while Events lost %s and Consulting %s. At the quarter gross margin of '
                '%.1f%%, revenue of approximately %s is required to cover overheads of %s, '
                'against the %s billed.'
                % (k(d['Total']['rev']), k(d['Total']['bud']), k(short),
                   short / d['Total']['bud'] * 100,
                   k(d['Integration']['bud'] - d['Integration']['rev']),
                   k(d['Consulting']['bud'] - d['Consulting']['rev']),
                   k(d['Support']['res']), k(abs(d['Events']['res'])),
                   k(abs(d['Consulting']['res'])), gpc * 100, k(breakeven), k(oh),
                   k(d['Total']['rev'])))
        setext(by, 'Text 53', para)
        s.notes_slide.notes_text_frame.text = para
        print('  slide %d rewritten to the quarter' % i)
        for name in DEPTS:
            e = d[name]
            print('     %-12s %8s  %4.0f%%  %8s after overhead'
                  % (name, k(e['rev']), e['gp'] / e['rev'] * 100, k(e['res'])))
        done = True
        break
    if not done:
        print('  executive summary not found')
    pres.save(out)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
