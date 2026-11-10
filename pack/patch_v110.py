"""Name the financial year the way the pack's own Setup sheet names it.

Setup!B7 is "Financial year (the year it ends)" and holds 2027, so this year
is FY27 and the three first quarters on the deck are Q1 FY25, Q1 FY26 and
Q1 FY27. Several charts were labelling them by the calendar year the quarter
falls in instead, Q1 2024, Q1 2025 and Q1 2026, so the same quarter was
called FY27 on one slide and 2026 on another.

Danica asked for FY27 across the whole deck. These are the label cells.
"""
import openpyxl

PACK = 'CTS Financial Controller Pack - v20.xlsx'
# sheet, cell, what it should say now, what it becomes
EDITS = [
    # the quarterly spend and profit block, derived from Setup!B7
    ('Xero History', 'A128', '="Q1 "&(Setup!$B$7-3)',
     '="Q1 FY"&(Setup!$B$7-2002)'),
    ('Xero History', 'A129', '="Q1 "&(Setup!$B$7-2)',
     '="Q1 FY"&(Setup!$B$7-2001)'),
    ('Xero History', 'A130', '="Q1 "&(Setup!$B$7-1)',
     '="Q1 FY"&(Setup!$B$7-2000)'),
    ('Xero History', 'A131', '="Budget Q1 "&(Setup!$B$7-1)',
     '="Budget Q1 FY"&(Setup!$B$7-2000)'),
    # the three year first quarter revenue block
    ('Xero History', 'A422', 'Q1 2024', 'Q1 FY25'),
    ('Xero History', 'A423', 'Q1 2025', 'Q1 FY26'),
    ('Xero History', 'A424', 'Q1 2026', 'Q1 FY27'),
    # the first quarter revenue chart on Cost Analysis, five years of labels
    ('Cost Analysis', 'A298', '="Q1 "&(2000+23-1)', '="Q1 FY"&23'),
    ('Cost Analysis', 'A299', '="Q1 "&(2000+24-1)', '="Q1 FY"&24'),
    ('Cost Analysis', 'A300', '="Q1 "&(2000+25-1)', '="Q1 FY"&25'),
    ('Cost Analysis', 'A301', '="Q1 "&(2000+26-1)', '="Q1 FY"&26'),
    ('Cost Analysis', 'A302', '="Q1 "&(2000+27-1)', '="Q1 FY"&27'),
]


def main():
    wb = openpyxl.load_workbook(PACK)
    done, skipped = 0, 0
    for sheet, ref, was, now in EDITS:
        cell = wb[sheet][ref]
        if cell.value == now:
            skipped += 1
            continue
        if cell.value != was:
            print('  STOPPED at %s!%s: expected %r, found %r'
                  % (sheet, ref, was, cell.value))
            return
        cell.value = now
        done += 1
        print('  %-14s %-5s %r -> %r' % (sheet, ref, was, now))
    wb.save(PACK)
    print('  %d label cells changed, %d already done' % (done, skipped))


if __name__ == '__main__':
    main()
