"""Push the full admin allocation through 'P&L by Dept'.

The tab allocated only the office staff pot, and it took that from
'Overheads MoM', which splits admin employment by labour hours rather than by
the Staff headcount key.  Every other admin category stayed 100 per cent in
General Admin.  So the monthly departmental P&L did not agree with 'Dept P&L'
or with the P&L analysis.

This repoints each overhead line at 'Overhead Split' sections 3 and 4, the same
subcategory SUMPRODUCT that 'Dept P&L' already uses, with the month taken from
each block's own date row instead of Setup!B5.  Direct departmental amounts
still come from the Engine.  General Admin carries the contra, so it nets to
nothing and the company total is unchanged.
"""
import re
import openpyxl

WB = 'pack/CTS Financial Controller Pack - v20.xlsx'
OS = "'Overhead Split'"
SUB = f'{OS}!$U$23:$U$86'
AMT = f'{OS}!$C$23:$N$86'
HDR = f'{OS}!$C$22:$N$22'

# Overhead lines in each block, in sheet order, with the subcategory each one
# draws its share of admin from.  'Other' rides with General & Administrative
# Exp: those are the admin accounts the analysis names but the pack's account
# list does not, and they are general administration by nature.
LINES = [
    (0, ['General & Administrative Exp', 'Other']),
    (1, ['Business Insurances']),
    (2, ['Advertising & Promotion Exp']),
    (3, ['Operating Expenses']),
    (6, ['Employment Expenses']),      # the office staff line
    (7, ['Credit Card fees paid']),
]

# name, date row, first overhead line row, key columns on the overhead split
BLOCKS = [
    ('Production',    6,  61, ['$Q$', '$R$']),   # Events + Video
    ('Consulting',   83, 138, ['$S$', '$T$']),   # Integration + Consulting
    ('Support',     160, 188, ['$P$']),
    ('General Admin', 210, 238, []),             # the contra
]

COLS = [openpyxl.utils.get_column_letter(c) for c in range(2, 14)]   # B..M


def alloc(col, date_row, keys, subs):
    """The department's share of admin for one month, one set of subcategories."""
    pick = f'INDEX({AMT},0,MATCH({col}${date_row},{HDR},0))'
    test = '+'.join(f'({SUB}="{s}")' for s in subs)
    if not keys:
        return f'-SUMPRODUCT(({test})*{pick})'
    key = '+'.join(f'{OS}!{k}23:{k}86' for k in keys)
    key = key if len(keys) == 1 else f'({key})'
    return f'SUMPRODUCT(({test})*{key}*{pick})'


wb = openpyxl.load_workbook(WB)
ws = wb['P&L by Dept']
changed = 0

for name, date_row, first, keys in BLOCKS:
    for offset, subs in LINES:
        row = first + offset
        for col in COLS:
            cell = ws[f'{col}{row}']
            old = cell.value
            if not isinstance(old, str) or not old.startswith('='):
                raise SystemExit(f'{name} {col}{row}: not a formula: {old!r}')
            share = alloc(col, date_row, keys, subs)
            if offset == 6:
                # the office staff line: allocation only, no direct amount
                new = f'={share}'
            else:
                m = re.fullmatch(r'=IF\((.+?)<=Setup!\$B\$5,(.+),(.+)\)', old)
                if not m:
                    raise SystemExit(f'{name} {col}{row}: unexpected shape: {old}')
                when, direct, forecast = m.groups()
                new = f'=IF({when}<=Setup!$B$5,{direct}+{share},{forecast})'
            cell.value = new
            changed += 1
    print(f'  {name:<14} rows {first}..{first + 7}')

wb.save(WB)
print(f'{changed} cells rewritten, saved')
