"""Read the controller pack and write data.json, the figures the deck's
commentary and speaker notes are built from.

Every number here is read from a cell in the pack. Nothing is typed in, so a
re-paste of the pack followed by a re-run moves the words on the slides as
well as the pictures.

The pack carries the previous month's cached results until Excel opens it, and
openpyxl only ever reads that cache. So the pack is put through LibreOffice
first, with the formulas stripped of their old answers, which forces every one
to be worked out again. That is the same route snip.py takes, which is why the
snips and these figures agree.

Run it from this folder:

    python3.13 extract.py

Anything it cannot read it says so and stops, rather than writing a figure
that looks right and is not.
"""
import datetime
import json
import os
import shutil
import subprocess
import sys
import tempfile

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
PACK = os.path.join(HERE, '..', '..', 'pack',
                    'CTS Financial Controller Pack - October 2026.xlsx')
OUT = os.path.join(HERE, 'data.json')


# ------------------------------------------------------------------ recalc
def recalculated(src):
    """A copy of the pack with every formula worked out afresh."""
    work = tempfile.mkdtemp(prefix='packcalc-')
    ind, outd = os.path.join(work, 'in'), os.path.join(work, 'out')
    os.makedirs(ind)
    os.makedirs(outd)
    mid = os.path.join(ind, 'pack.xlsx')
    print('  loading the pack', flush=True)
    openpyxl.load_workbook(src).save(mid)      # formulas, no cached answers
    print('  recalculating, this takes a few minutes', flush=True)
    subprocess.run(['soffice', '--headless', '--norestore', '--convert-to',
                    'xlsx', '--outdir', outd, mid],
                   capture_output=True, text=True, timeout=3600,
                   env=dict(os.environ, HOME=work))
    made = os.path.join(outd, 'pack.xlsx')
    if not os.path.exists(made):
        raise SystemExit('LibreOffice produced nothing')
    keep = os.path.join(tempfile.mkdtemp(prefix='packcalc-out-'), 'pack.xlsx')
    shutil.copy(made, keep)
    shutil.rmtree(work, ignore_errors=True)
    return keep


# ------------------------------------------------------------- cell access
class Sheet:
    def __init__(self, wb, name, rows):
        self.name = name
        self.rows = list(wb[name].iter_rows(min_row=1, max_row=rows))

    def __call__(self, r, c):
        try:
            return self.rows[r - 1][c - 1].value
        except IndexError:
            return None

    def num(self, r, c):
        v = self(r, c)
        if v is None:
            return 0.0
        if not isinstance(v, (int, float)):
            raise SystemExit('%s!%s%d is %r, expected a number'
                             % (self.name, chr(64 + c), r, v))
        return float(v)

    def soft(self, r, c):
        """Like num, but a heading where a figure would be reads as nothing."""
        v = self(r, c)
        return float(v) if isinstance(v, (int, float)) else 0.0

    def text(self, r, c):
        v = self(r, c)
        return '' if v is None else str(v).strip()

    def month_col(self, header_row, month, first=2, last=30):
        """The column whose header on header_row is this month."""
        for c in range(first, last + 1):
            v = self(header_row, c)
            if isinstance(v, datetime.datetime) and \
                    (v.year, v.month) == (month.year, month.month):
                return c
        raise SystemExit('%s row %d has no column for %s'
                         % (self.name, header_row, month.strftime('%b %Y')))


def main():
    # A pack that has already been through LibreOffice can be passed in, so a
    # change to the mapping can be checked without waiting for the recalc.
    path = sys.argv[1] if len(sys.argv) > 1 else recalculated(PACK)
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    S = Sheet(wb, 'Summary', 190)
    A26 = Sheet(wb, 'Actual FY26', 190)
    PL27 = Sheet(wb, 'P&L FY27', 190)
    BUD = Sheet(wb, 'Budget FY27', 190)
    OS = Sheet(wb, 'Overhead Split', 160)
    OM = Sheet(wb, 'Overheads MoM', 170)
    CA = Sheet(wb, 'Cost Analysis', 80)
    DSO = Sheet(wb, 'DSO', 45)
    T10 = Sheet(wb, 'Top 10 Clients', 20)
    UT = Sheet(wb, 'Utilisation', 65)

    month = S(3, 2)
    if not isinstance(month, datetime.datetime):
        raise SystemExit('Summary!B3 is %r, expected the reporting month' % month)
    ly = datetime.datetime(month.year - 1, month.month, 1)
    print('  reporting month:', month.strftime('%B %Y'), flush=True)

    D = {}

    # ----------------------------------------------------------- the P&L
    # Column D is the reporting month, E the year to date, G the year to date
    # budget and L the month's budget. Last year comes off Actual FY26, which
    # has a column for each month and its own year to date in column O.
    PL_ROWS = {'revenue': 6, 'other_income': 57, 'cos': 59, 'overheads': 164,
               'gross_profit': 173, 'gp_pct': 174, 'da': 178,
               'ebitda': 175, 'ebitda_pct': 176,
               'net_profit': 181, 'np_pct': 182}
    ly_col = A26.month_col(5, ly)
    D['pl'] = {k: {'month': S.num(r, 4), 'month_budget': S.num(r, 12),
                   'month_ly': A26.num(r, ly_col),
                   'ytd': S.num(r, 5), 'ytd_budget': S.num(r, 7),
                   'ytd_ly': A26.num(r, 15)}
               for k, r in PL_ROWS.items()}

    # ---------------------------------------------------- the departments
    # Overhead Split rows 143 to 147 are the department result for the month
    # on Setup. Revenue and cost of sales by department sit on the Summary.
    REV_ROW = {'Support': 41, 'Events': 8, 'Video': 16,
               'Integration': 25, 'Consulting': 33}
    COS_ROW = {'Support': 130, 'Events': 61, 'Video': 78,
               'Integration': 96, 'Consulting': 113}
    own_col = OS.month_col(100, month)
    depts = []
    for i, r in enumerate(range(143, 148)):
        name = OS.text(r, 1)
        if name not in REV_ROW:
            raise SystemExit('Overhead Split row %d reads %r, '
                             'not a department I know' % (r, name))
        rr, cr = REV_ROW[name], COS_ROW[name]
        own = OS.num(101 + i, own_col)
        allo = OS.num(r, 7)
        depts.append({
            'name': name,
            'revenue': OS.num(r, 2), 'budget': OS.num(r, 3),
            'ly': A26.num(rr, ly_col),
            'cos': S.num(cr, 4), 'gross_profit': OS.num(r, 5),
            'gm_pct': OS.num(r, 6),
            'own_overhead': own, 'alloc_overhead': allo - own,
            'all_overhead': allo, 'result': OS.num(r, 8),
            'ytd_revenue': S.num(rr, 5), 'ytd_budget': S.num(rr, 7),
            'ytd_cos': S.num(cr, 5), 'cos_budget': S.num(cr, 12),
        })
    D['dept'] = depts

    # ------------------------------------------------- people and overhead
    D['staff'] = [{'name': OM.text(r, 1), 'departmental': OM.num(r, 2),
                   'office': OM.num(r, 3),
                   'total': OM.num(r, 2) + OM.num(r, 3)}
                  for r in range(162, 167)]

    fte_col = OS.month_col(112, month)
    D['fte'] = [{'name': OS.text(113 + i, 1), 'hours': OS.num(113 + i, fte_col),
                 'share': OS.num(121 + i, fte_col)} for i in range(5)]
    D['annual_hours'] = OS.num(119, fte_col)

    # Utilisation hours are pasted on to the Utilisation tab each month.
    # Each department is a block of five rows: billable, non-billable, leave,
    # billable share, FTE. Worked hours are rows 58 to 62.
    ut_col = UT.month_col(8, month)
    util = []
    for i in range(5):
        b = 9 + i * 6
        util.append({'name': UT.text(b, 2), 'chargeable': UT.num(b, ut_col),
                     'non_chargeable': UT.num(b + 1, ut_col),
                     'leave': UT.num(b + 2, ut_col),
                     'util_pct': UT.num(b + 3, ut_col),
                     'fte': UT.num(b + 4, ut_col),
                     'worked_cc': UT.num(58 + i, ut_col)})
    D['util'] = util
    D['working_hours_month'] = UT.num(6, UT.month_col(4, month))
    if not any(u['chargeable'] for u in util):
        print('  WARNING: no utilisation hours for %s on the Utilisation tab'
              % month.strftime('%B %Y'), flush=True)

    # ---------------------------------------------------------- customers
    D['top10'] = [{'rank': int(T10.num(r, 1)), 'name': T10.text(r, 6),
                   'month': T10.num(r, 3), 'ytd': T10.num(r, 4),
                   'share': T10.num(r, 5)} for r in range(7, 17)]
    # C18 and C19 on that tab are wired to a fixed column on P&L FY27 rather
    # than to the reporting month, so "everyone else" for the month is worked
    # out here from the month's total income instead of read from C18.
    month_income = S.num(6, 4)
    D['top10_total'] = {
        'top_ten_month': T10.num(17, 3), 'top_ten_ytd': T10.num(17, 4),
        'top_ten_share': T10.num(17, 5),
        'rest_month': month_income - T10.num(17, 3),
        'rest_ytd': T10.num(18, 4), 'rest_share': T10.num(18, 5),
    }

    # ------------------------------------------------------ cash and debt
    D['dso'] = {'receivables': DSO.num(6, 2), 'revenue_3m': DSO.num(7, 2),
                'days_in': DSO.num(8, 2), 'days': DSO.num(10, 2),
                'target': DSO.num(11, 2), 'above': DSO.num(12, 2),
                'locked': DSO.num(13, 2), 'at_target': DSO.num(14, 2),
                'overdue': DSO.num(26, 2), 'overdue_share': DSO.num(26, 3)}
    D['ageing'] = [{'bucket': DSO.text(r, 1), 'amount': DSO.num(r, 2),
                    'share': DSO.num(r, 3)} for r in range(19, 25)]
    D['debtors'] = [{'name': DSO.text(r, 2), 'amount': DSO.num(r, 3),
                     'share': DSO.num(r, 4)} for r in range(31, 41)]

    # ----------------------------------------------------- what we spend
    D['other_opex'] = [{'label': CA.text(r, 1), 'amount': CA.num(r, 2),
                        'share': CA.num(r, 3)} for r in range(6, 15)]
    D['cost_dist'] = [{'label': CA.text(r, 1), 'amount': CA.num(r, 2),
                       'share': CA.num(r, 3)} for r in range(49, 59)]
    # Row 63 is the heading, kept so the block matches the tab row for row.
    # Amount is the direct column and share the indirect one.
    D['labour_mix'] = [{'label': CA.text(r, 1), 'amount': CA.soft(r, 2),
                        'share': CA.soft(r, 3)} for r in range(63, 79)]

    # ------------------------------------------------- budget and months
    D['budget_actual'] = [
        {'label': lab, 'amount': S.num(r, 5), 'share': S.num(r, 7)}
        for lab, r in [('Total Income', 6), ('Cost of Sales', 59),
                       ('Gross Profit', 173), ('Overheads', 164),
                       ('EBITDA', 175), ('Net Profit', 181)]]

    months = [PL27(5, c) for c in range(2, 14)]
    D['income_by_month'] = [{'month': str(months[i]),
                             'actual': PL27.num(6, 2 + i),
                             'budget': BUD.num(6, 2 + i)} for i in range(12)]
    D['profit_by_month'] = [{'month': str(months[i]),
                             'a': PL27.num(173, 2 + i),
                             'b': PL27.num(175, 2 + i),
                             'c': PL27.num(181, 2 + i)} for i in range(12)]

    # ------------------------------------------------- the first quarter
    # Q1 is fixed to July, August and September on its own tab, so the
    # quarterly slides keep reading Q1 after the pack moves on a month.
    Q1 = Sheet(wb, 'Q1 Company Summary', 190)
    q1b = lambda r: sum(BUD.num(r, c) for c in range(2, 5))   # Jul, Aug, Sep
    D['q1'] = {k: {'actual': Q1.num(r, 2), 'ly': Q1.num(r, 3),
                   'budget': q1b(r)}
               for k, r in PL_ROWS.items() if k in
               ('revenue', 'cos', 'overheads', 'gross_profit', 'ebitda',
                'net_profit')}

    # Full year, for the slide that sets the quarter against the year.
    D['fy'] = {k: {'budget': BUD.num(r, 14), 'ly': A26.num(r, 14)}
               for k, r in PL_ROWS.items() if k in
               ('revenue', 'gross_profit', 'ebitda', 'net_profit')}

    D['setup_month'] = month.strftime('%B %Y')
    D['months'] = {'short': month.strftime('%b %y').replace(' 2', ' 2'),
                   'ly_short': ly.strftime('%b %y'),
                   'name': month.strftime('%B'),
                   'ly_name': ly.strftime('%B'),
                   'period': int(((month.month - 7) % 12) + 1),
                   'end': (datetime.datetime(month.year + (month.month == 12),
                                             month.month % 12 + 1, 1)
                           - datetime.timedelta(days=1)).strftime('%d %B %Y')
                           .lstrip('0')}

    with open(OUT, 'w') as f:
        json.dump(D, f, indent=1)
    print('  wrote', OUT, flush=True)
    print('  revenue %s, EBITDA %s, net profit %s'
          % (round(D['pl']['revenue']['month'], 2),
             round(D['pl']['ebitda']['month'], 2),
             round(D['pl']['net_profit']['month'], 2)), flush=True)
    if len(sys.argv) <= 1:
        shutil.rmtree(os.path.dirname(path), ignore_errors=True)


if __name__ == '__main__':
    main()
