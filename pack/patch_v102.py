"""Re-split the Consulting group budget between Integration and Consulting.

Danica moved the Aware Super work from CONS to INT in her budget model. That
model stops at three departments and never splits Integration from Consulting,
so the pack has always done it with a flat ratio. The ratio was 72 / 28.

Her sales forecast, weighted, now carries 535,720.23 against INT and
104,157.10 against CONS for the full year. That is the ratio applied here, and
Danica approved it.

The group totals do not move. Only the split inside them. Income has one live
line on each side, so it is set outright. Cost has many, so each block is
scaled by one factor, which keeps the mix of accounts inside it unchanged.

The group rows are formulas that sum their detail, so only the detail is
written and the totals follow.
"""
import sys
import openpyxl

INT_FC, CONS_FC = 535720.23, 104157.10        # weighted, from CONS Sales Forecast
R_INT = INT_FC / (INT_FC + CONS_FC)
R_CONS = 1 - R_INT

INC = {'group': (25, 33), 'int_live': 27, 'cons_live': 35,
       'int_rows': range(26, 33), 'cons_rows': range(34, 41)}
COS = {'group': (96, 113),
       'int_rows': range(97, 113), 'cons_rows': range(114, 130)}
MONTHS = range(2, 14)


def main(path):
    wb = openpyxl.load_workbook(path)
    ws = wb['Budget FY27']
    print('new split: Integration %.4f%%, Consulting %.4f%%' % (R_INT * 100, R_CONS * 100))

    # income: every line but the one live row on each side must be empty,
    # otherwise setting the live row outright would lose something
    for rows, live in ((INC['int_rows'], INC['int_live']),
                       (INC['cons_rows'], INC['cons_live'])):
        for r in rows:
            if r == live:
                continue
            for c in MONTHS:
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)) and abs(v) > 0.005:
                    print('STOPPED: income row %d column %d holds %.2f' % (r, c, v))
                    return

    moved = 0.0
    for c in MONTHS:
        i_now = ws.cell(INC['int_live'], c).value or 0
        c_now = ws.cell(INC['cons_live'], c).value or 0
        grp = i_now + c_now
        ws.cell(INC['int_live'], c).value = grp * R_INT
        ws.cell(INC['cons_live'], c).value = grp * R_CONS
        moved += grp * R_INT - i_now

        # the group cost is read once, before either block is touched. Reading
        # it inside the loop uses the already scaled Integration rows on the
        # second pass and quietly moves the group total.
        int_now = sum(ws.cell(r, c).value or 0 for r in COS['int_rows'])
        cons_now = sum(ws.cell(r, c).value or 0 for r in COS['cons_rows'])
        grp_cost = int_now + cons_now
        for rows, cur, share in ((COS['int_rows'], int_now, R_INT),
                                 (COS['cons_rows'], cons_now, R_CONS)):
            if abs(cur) < 0.005:
                continue
            f = grp_cost * share / cur
            for r in rows:
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)) and v:
                    ws.cell(r, c).value = v * f

    print('income moved to Integration across the year: %.2f' % moved)
    wb.save(path)
    print('saved')


if __name__ == '__main__':
    main(sys.argv[1])
