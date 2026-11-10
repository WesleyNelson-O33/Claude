"""A profit and loss slide for each reporting group, three months and the
quarter side by side.

Danica asked for the Production block from the P&L by Dept sheet on its own
slide, with July, August, September and the year to date together. The other
two groups get the same, because a section with only Production in it reads
as a gap.

These tables are close to square, so they take the full slide height with no
commentary strip, the same as the quarter summary. The commentary goes into
the speaker notes.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from addslide import add

SNIP = '/home/user/Claude/deck/snips/'

GROUPS = [
    ('pl_prd', 'Production Profit And Loss, First Quarter FY2027', [
        'Production billed $460,391 for the quarter and kept $272,628 at the gross line, 59.2%.',
        'August carried it. $268,942 of the quarter\'s revenue landed in that one month.',
        'July and September billed $110,843 and $80,607, both loss making months.',
        '$325,700 of overhead turns a $272,628 gross profit into a $52,081 net loss.',
        'Events is almost all of it. Video billed $40,270 across the three months.',
    ]),
    ('pl_cons', 'Consulting Profit And Loss, First Quarter FY2027', [
        'Consulting billed $445,481 for the quarter but kept only $48,782, an 11.0% gross margin.',
        'September is the month. $238,389 of the quarter\'s revenue arrived in it.',
        'Integration is the volume at $441,414. The Consulting department billed $4,068.',
        '$75,653 of overhead against $48,782 of gross profit leaves a $26,777 net loss.',
        'This is turnover without margin rather than a cost problem.',
    ]),
    ('pl_sup', 'Support Profit And Loss, First Quarter FY2027', [
        'Support billed $718,587 for the quarter at a 43.3% gross margin.',
        'It is the steadiest book in the business, between $235k and $243k every month.',
        '$240,184 of overhead against $310,919 of gross profit leaves $73,050 of net profit.',
        'A 10.2% net margin, and the only reporting group above the line.',
        'It is covering the losses in the other two.',
    ]),
]

NOTES = {
    'pl_prd': (
        'Production for the quarter, month by month and in total.\n\n'
        'Revenue was $460,391 and gross profit $272,628, a margin of 59.2 per cent, '
        'which is the best gross margin of the three groups.\n\n'
        'The shape is the problem, not the margin. August billed $268,942 and made '
        '$44,386. July billed $110,843 and lost $43,623. September billed $80,607 and '
        'lost $52,844. One month in three carried it.\n\n'
        'Overhead is $325,700 for the quarter and does not move with revenue, so in the '
        'two quiet months it swamps the gross profit. That is what turns a $272,628 '
        'gross profit into a $52,081 net loss.\n\n'
        'Events is almost the whole group. Video billed $40,270 across the three months.'),
    'pl_cons': (
        'Consulting for the quarter, month by month and in total.\n\n'
        'Revenue was $445,481, within $15,000 of Production, but gross profit was only '
        '$48,782. That is an 11.0 per cent margin against Production\'s 59.2.\n\n'
        'September is where it lands. $238,389 of the quarter arrived in that month on '
        'project completion.\n\n'
        'Integration is the volume at $441,414. The Consulting department itself billed '
        '$4,068 against a budget of $236,707.\n\n'
        'Overhead of $75,653 against $48,782 of gross profit gives a $26,777 net loss. '
        'The overhead is modest. The issue is that the work is being sold at a margin '
        'that cannot carry it.'),
    'pl_sup': (
        'Support for the quarter, month by month and in total.\n\n'
        'Revenue was $718,587 at a 43.3 per cent gross margin, giving $310,919 of gross '
        'profit.\n\n'
        'It is the steadiest book in the business. $235,403 in July, $240,236 in August '
        'and $242,948 in September. That is what a contract base should look like.\n\n'
        'Overhead of $240,184 leaves $73,050 of net profit, a 10.2 per cent net margin.\n\n'
        'Support is the only reporting group above the line this quarter, and it is '
        'covering the losses in the other two.'),
}


def main(src, out, at):
    pres = Presentation(src)
    before = len(pres.slides)
    tpl = next(i for i, sl in enumerate(pres.slides)
               if {'commentary panel', 'Text 4', 'Text 5'} <= {sh.name for sh in sl.shapes}
               and any('PICTURE' in str(sh.shape_type) and sh.name.startswith('snip ')
                       for sh in sl.shapes))
    at = int(at)
    for n, (snip, title, bullets) in enumerate(GROUPS):
        add(pres, tpl, at + n, 'DEPARTMENTAL PERFORMANCE', title,
            SNIP + snip + '.png', bullets, NOTES[snip], 'snip ' + snip)
        print('  %-8s at page %d  %s' % (snip, at + n + 1, title))
    pres.save(out)
    print('%d slides in, %d out' % (before, len(pres.slides)))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3])
