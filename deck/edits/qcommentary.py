"""Rewrite the commentary on every quarterly slide, from the pack.

The figures in these slides were typed in by hand when each slide was built.
When the pack was re-pasted from Xero they went stale, silently, which is
exactly the failure this avoids. Nothing here is typed. Every number is read
from the controller pack at the moment the slide is written.

Three kinds of slide:
  the executive summary, which is cards and a strip
  the ordinary chart pages, which have a commentary strip
  the full height tables, which have no strip, so the text goes in the notes
"""
import copy, sys, os
import openpyxl
from pptx import Presentation
from pptx.util import Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from addslide import balance, write, SIZE

PACK = '/home/user/Claude/pack/CTS Financial Controller Pack - v20.xlsx'
INC = {'Support': 41, 'Events': 8, 'Video': 16,
       'Integration': 25, 'Consulting': 33, 'Total': 6}
COS = {'Support': 130, 'Events': 61, 'Video': 78,
       'Integration': 96, 'Consulting': 113, 'Total': 59}
GROUP_COL = {'Production': 2, 'Consulting': 3, 'Support': 4, 'CTS': 6}
GROUP_ROW = {'income': 272, 'cos': 274, 'oh': 275, 'gp': 277,
             'gpc': 278, 'ebitda': 279, 'net': 285, 'netc': 286}


def k(v):
    n = round(v / 1000)
    return ('$(%s)k' % format(abs(n), ',')) if n < 0 else ('$%sk' % format(n, ','))


def d(v):
    return ('$(%s)' % format(abs(round(v)), ',')) if v < 0 else ('$%s' % format(round(v), ','))


def figures():
    wb = openpyxl.load_workbook(PACK, data_only=True)
    sm, pd = wb['Summary'], wb['P&L by Dept']
    f = {'dept': {}, 'grp': {}}
    for name in INC:
        rev = sm.cell(INC[name], 5).value
        bud = sm.cell(INC[name], 7).value or 0
        gp = rev + sm.cell(COS[name], 5).value
        f['dept'][name] = {'rev': rev, 'bud': bud, 'gp': gp,
                           'gm': gp / rev * 100 if rev else 0}
    for name, col in GROUP_COL.items():
        f['grp'][name] = {key: pd.cell(row, col).value
                          for key, row in GROUP_ROW.items()}
    f['oh'] = abs(sm.cell(164, 5).value)
    f['oh_b'] = abs(sm.cell(164, 7).value)
    f['gp'] = sm.cell(173, 5).value
    f['gpc'] = sm.cell(174, 5).value
    f['gpc_b'] = sm.cell(174, 7).value
    f['eb'] = sm.cell(175, 5).value
    f['eb_b'] = sm.cell(175, 7).value
    f['net'] = sm.cell(181, 5).value
    f['net_b'] = sm.cell(181, 7).value
    f['breakeven'] = f['oh'] / f['gpc'] if f['gpc'] else 0
    return f


def build_text(f):
    t, g = f['dept'], f['grp']
    short = t['Total']['bud'] - t['Total']['rev']
    two_down = abs(g['Production']['net']) + abs(g['Consulting']['net'])

    summary = [
        'Revenue for the quarter was %s against a budget of %s, short by %s or %.0f%%.'
        % (d(t['Total']['rev']), d(t['Total']['bud']), d(short), short / t['Total']['bud'] * 100),
        'Gross profit was %s, a margin of %.1f%% against the %.1f%% budgeted.'
        % (d(f['gp']), f['gpc'] * 100, f['gpc_b'] * 100),
        'Overheads were %s, which is %s under budget.' % (d(f['oh']), d(f['oh_b'] - f['oh'])),
        'EBITDA was %s against a budget of %s.' % (d(f['eb']), d(f['eb_b'])),
        'The quarter closed as a net loss of %s, against a budgeted profit of %s.'
        % (d(abs(f['net'])), d(f['net_b'])),
    ]
    group = [
        'Revenue for the quarter was %s and gross profit %s, a margin of %.1f%%.'
        % (d(t['Total']['rev']), d(f['gp']), f['gpc'] * 100),
        'Support is the only reporting group above the line, earning %s on a %.1f%% net margin.'
        % (d(g['Support']['net']), g['Support']['netc'] * 100),
        'Production lost %s across the three months and Consulting lost %s.'
        % (d(abs(g['Production']['net'])), d(abs(g['Consulting']['net']))),
        'Group EBITDA was %s and the quarter closed as a net loss of %s.'
        % (d(f['eb']), d(abs(f['net']))),
        'Support carries the quarter. The other two groups are %s down between them.'
        % d(two_down),
    ]
    pl = [
        'Support is the only group that earns its overhead back, at %s net on a %.1f%% margin.'
        % (d(g['Support']['net']), g['Support']['netc'] * 100),
        'Production holds the best gross margin at %.1f%%, then loses it to %s of overhead.'
        % (g['Production']['gpc'] * 100, d(abs(g['Production']['oh']))),
        'Consulting keeps only %.1f%% at the gross line, the weakest of the three.'
        % (g['Consulting']['gpc'] * 100),
        'Production and Consulting are %s down between them for the quarter.' % d(two_down),
        'Group EBITDA was %s and the quarter closed as a net loss of %s.'
        % (d(f['eb']), d(abs(f['net']))),
    ]
    ytd = [
        'Support billed %s year to date, the largest of the five departments.' % d(t['Support']['rev']),
        'Integration finished the quarter at %s and Events at %s.'
        % (d(t['Integration']['rev']), d(t['Events']['rev'])),
        'Integration started the year behind Events and overtook it on project completion.',
        'Video billed %s year to date and Consulting %s.'
        % (d(t['Video']['rev']), d(t['Consulting']['rev'])),
        'The five together make the %s of year to date revenue.' % d(t['Total']['rev']),
    ]
    return {'summary': summary, 'group': group, 'pl': pl, 'ytd': ytd}


def grp_notes(f, name, label):
    g = f['grp'][name]
    return ('%s for the quarter, month by month and in total.\n\n'
            'Revenue was %s and gross profit %s, a margin of %.1f per cent.\n\n'
            'Overhead of %s leaves %s, a net margin of %.1f per cent.\n\n'
            'Figures are read from the controller pack and tie to Xero.'
            % (label, d(g['income']), d(g['gp']), g['gpc'] * 100,
               d(abs(g['oh'])), d(g['net']), g['netc'] * 100))


TARGETS = {
    'results by reporting group, first quarter fy2027': 'group',
    'profit and loss by reporting group, first quarter fy2027': 'pl',
    'year to date revenue by department': 'ytd',
}
NOTE_ONLY = {
    'profit and loss summary, first quarter fy2027': 'summary',
    'production profit and loss, first quarter fy2027': ('Production', 'Production'),
    'consulting profit and loss, first quarter fy2027': ('Consulting', 'Consulting'),
    'support profit and loss, first quarter fy2027': ('Support', 'Support'),
}


def title_of(slide):
    for sh in slide.shapes:
        if sh.name == 'Text 1' and sh.has_text_frame:
            return sh.text_frame.text.strip().split('\n')[0]
    return ''


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


def main(src, out):
    f = figures()
    text = build_text(f)
    pres = Presentation(src)
    done = 0
    for i, s in enumerate(pres.slides, 1):
        t = title_of(s).lower()
        by = {sh.name: sh for sh in s.shapes}

        if t == 'performance summary, first quarter fy2027':
            for name, (big, marg, third, budg) in CARDS.items():
                e = f['dept'][name]
                setext(by, big, k(e['rev']))
                setext(by, marg, '%.0f%% gross margin' % e['gm'])
                setext(by, third, '%s gross profit' % k(e['gp']))
                diff = e['rev'] - e['bud']
                setext(by, budg, ('Ahead of budget by %s' if diff >= 0 else 'Behind budget by %s') % k(abs(diff)))
            vals = [(k(f['dept']['Total']['rev']), 'budget %s' % k(f['dept']['Total']['bud'])),
                    ('%.1f%%' % (f['gpc'] * 100), 'budget %.1f%%' % (f['gpc_b'] * 100)),
                    (k(f['oh']), 'budget %s' % k(f['oh_b'])),
                    (k(f['eb']), 'budget %s' % k(f['eb_b'])),
                    (k(f['breakeven']), 'running at %s' % k(f['dept']['Total']['rev']))]
            for (v, sub), (vb, sb) in zip(vals, STRIP):
                setext(by, vb, v); setext(by, sb, sub)
            g = f['grp']
            short = f['dept']['Total']['bud'] - f['dept']['Total']['rev']
            para = ('The quarter in summary. Revenue was %s against a budget of %s, short by %s '
                    'or %.0f%%. Almost all of the shortfall is Consulting at %s and Integration '
                    'at %s. Support is the only reporting group that covered its allocated '
                    'overhead, earning %s, while Production lost %s and Consulting %s. At the '
                    'quarter gross margin of %.1f%%, revenue of approximately %s is required to '
                    'cover overheads of %s, against the %s billed.'
                    % (k(f['dept']['Total']['rev']), k(f['dept']['Total']['bud']), k(short),
                       short / f['dept']['Total']['bud'] * 100,
                       k(f['dept']['Consulting']['bud'] - f['dept']['Consulting']['rev']),
                       k(f['dept']['Integration']['bud'] - f['dept']['Integration']['rev']),
                       k(g['Support']['net']), k(abs(g['Production']['net'])),
                       k(abs(g['Consulting']['net'])), f['gpc'] * 100, k(f['breakeven']),
                       k(f['oh']), k(f['dept']['Total']['rev'])))
            setext(by, 'Text 53', para)
            s.notes_slide.notes_text_frame.text = para
            print('  %-3d executive summary, cards and strip rewritten' % i)
            done += 1
            continue

        if t in TARGETS:
            items = text[TARGETS[t]]
            cols = list(balance(items, 4.05))
            for name, part in (('Text 5', cols[0]), ('Text 5b', cols[1])):
                box = by.get(name)
                if box is None or not part:
                    continue
                write(box, part)
                for para in box.text_frame.paragraphs:
                    for r in para.runs:
                        r.font.size = Pt(SIZE)
            print('  %-3d %s' % (i, title_of(s)))
            done += 1
            continue

        if t in NOTE_ONLY:
            spec = NOTE_ONLY[t]
            body = ('\n'.join(text['summary']) if spec == 'summary'
                    else grp_notes(f, spec[0], spec[1]))
            s.notes_slide.notes_text_frame.text = body
            print('  %-3d %s  (notes)' % (i, title_of(s)))
            done += 1

    pres.save(out)
    print('%d slides rewritten from the pack' % done)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
