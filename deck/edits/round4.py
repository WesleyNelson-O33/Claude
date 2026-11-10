"""Round four.

1. Name the two budget misses on the revenue slide (the executive
   summary paragraph is rewritten by q1exec_inplace.py, which reads the pack) Danica gave me:
     PwC budgeted 336k in August, did not materialise and will not.
     Bankwest budgeted in September, moved to the second quarter.
   Both tie to her budget model. PwC sits in the CONS sales forecast tagged
   INT at 335,984 in August. Bankwest sits at 48,984 in September.

2. Turn the cover and the section dividers from monthly to quarterly. The
   deck is the Q1 FY2027 review, the cover still said MONTHLY BUSINESS
   REVIEW, September 2026.
"""
import sys, os
import openpyxl
from pptx import Presentation

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from herdeck import title_of, write_panel, fill_box
from pptx.util import Inches

PACK = '/home/user/Claude/pack/CTS Financial Controller Pack - v20.xlsx'
PWC, BANKWEST = 335984, 48984

# cover and divider wording. left is matched whole, right replaces it.
RETITLE = {
    'MONTHLY BUSINESS REVIEW': 'QUARTERLY BUSINESS REVIEW',
    'September 2026': 'Q1 FY2027',
    "Every department's result for September 2026 and how it was made":
        "Every department's result for Q1 FY2027 and how it was made",
    'The P&L for September 2026 and the year to date':
        'The P&L for Q1 FY2027 and the year to date',
}


def d(v):
    return '$%s' % format(round(v), ',')


def k(v):
    return '$%sk' % format(round(v / 1000), ',')


def retitle(pres):
    hits = []
    for i, s in enumerate(pres.slides, 1):
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            for para in sh.text_frame.paragraphs:
                whole = ''.join(r.text for r in para.runs)
                new = RETITLE.get(whole.strip())
                if not new or not para.runs:
                    continue
                # keep the first run's formatting, empty the rest
                para.runs[0].text = new
                for extra in para.runs[1:]:
                    extra.text = ''
                hits.append('%d %s -> %s' % (i, whole.strip()[:40], new))
    cp = pres.core_properties
    cp.title = 'CTS Quarterly Business Review - Q1 FY2027'
    cp.subject = 'Quarterly Business Review'
    return hits


def main(src, out):
    wb = openpyxl.load_workbook(PACK, data_only=True)
    sm, xh = wb['Summary'], wb['Xero History']
    ytd, bud = sm.cell(6, 5).value, sm.cell(6, 7).value
    short = bud - ytd
    fy25_q1, fy26_q1 = xh.cell(60, 2).value, xh.cell(60, 3).value

    pres = Presentation(src)
    done = []
    for i, s in enumerate(pres.slides, 1):
        t = title_of(s).lower()

        if t.startswith('year to date revenue'):
            by = {sh.name: sh for sh in s.shapes}
            left = [
                'Year to date revenue is %s against a budget of %s.' % (d(ytd), d(bud)),
                'That is %s or %.0f%% behind.' % (d(short), short / bud * 100),
                'PwC was budgeted at %s in August. That work will not now happen.' % k(PWC),
            ]
            right = [
                'Bankwest was budgeted at %s in September and has moved into Q2.' % k(BANKWEST),
                'Revenue is still ahead of both completed years, %s in FY25 and %s in FY26.'
                % (d(fy25_q1), d(fy26_q1)),
                'The dashed line is the full year plan. Nine months are still to come.',
            ]
            fill_box(by['Text 5'], left)
            fill_box(by['Text 5b'], right)
            # the heading box ends at 4.29. the body started at 4.23 and ran
            # into it, and six lines spilled past the panel bottom at 5.06.
            for nm in ('Text 5', 'Text 5b'):
                by[nm].top = Inches(4.30)
                by[nm].height = Inches(0.74)
            done.append('%d year to date revenue' % i)

    for h in retitle(pres):
        done.append(h)
    pres.save(out)
    for x in done:
        print('  %s' % x)
    print('%d slides out' % len(pres.slides))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
