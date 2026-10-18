"""
Strip the explanatory notes out of the workbook.

Every sheet carried a sentence or two under its headings explaining what the
block was for, plus helper text beside the typed cells on Setup and long
"Check, ... should be nil" labels. They clutter the sheets and they show up in
the deck snips, so they go.

What is kept:
  README, which is the one sheet whose whole purpose is the notes
  the pasted Xero sheets, which are data, not commentary
  every check row, shortened to the word Check so the control survives
  headings, titles and column labels

A cell is treated as a note when it sits in the label columns, runs to six
words or more, and reads as a sentence rather than an account name. Account and
customer names are long but carry almost no stopwords, which is what separates
them.
"""
import openpyxl, re, sys

WB = sys.argv[1] if len(sys.argv) > 1 else 'CTS Financial Controller Pack - v20.xlsx'
KEEP_SHEETS = {'README', 'PL_Paste', 'BS_Paste', 'AR_Paste', 'P&L Analysis'}
LABEL_COLS = {1, 2, 3, 4}            # notes live in the label columns, never in the figures
STOPWORDS = {
    'the', 'a', 'an', 'is', 'are', 'was', 'were', 'be', 'so', 'which', 'that',
    'this', 'these', 'it', 'its', 'you', 'your', 'and', 'of', 'to', 'from',
    'on', 'in', 'into', 'every', 'with', 'not', 'no', 'if', 'then', 'but',
    'has', 'have', 'had', 'for', 'at', 'as', 'by', 'or', 'their', 'there',
    'what', 'when', 'where', 'how', 'why', 'do', 'does', 'did', 'can', 'will',
}


def is_note(text):
    """Reads as a sentence rather than a name. Account and customer names are
    long but carry almost no stopwords, which is what separates them."""
    t = text.strip()
    if len(t) < 30 or t.startswith('='):
        return False
    if t.isupper():                                  # section heading
        return False
    if re.match(r'^\d+\.', t):                       # numbered section heading
        return False
    words = re.findall(r"[A-Za-z']+", t)
    if len(words) < 6:
        return False
    hits = sum(1 for w in words if w.lower() in STOPWORDS)
    return hits >= 3


def row_has_data(ws, row):
    """A label has figures beside it. A note sits in a row of its own."""
    import datetime as _dt
    for c in ws[row]:
        if isinstance(c.value, (int, float, _dt.date, _dt.datetime)):
            return True
        if isinstance(c.value, str) and c.value.startswith('='):
            return True
    return False


def main():
    wb = openpyxl.load_workbook(WB)
    cleared, shortened, titles = [], [], []
    for ws in wb.worksheets:
        if ws.title in KEEP_SHEETS:
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str) or c.column not in LABEL_COLS:
                    continue
                t = v.strip()
                # A sheet title carrying a clause: keep the name, drop the clause.
                if c.row == 1 and c.column == 1 and ' - ' in t.replace('  -  ', ' - '):
                    name = re.split(r'\s+-\s+', t)[0].strip()
                    if name and name != t:
                        c.value = name
                        titles.append((ws.title, c.coordinate, t))
                        continue
                # A check row keeps its control, loses its clause. A label such
                # as "Check against the balance sheet" stays; anything carrying
                # a comma or a dash is a sentence and goes down to one word.
                if re.match(r'^check\b', t, re.I) and re.search(r'[,\u2013-]', t):
                    if t != 'Check':
                        c.value = 'Check'
                        shortened.append((ws.title, c.coordinate, t))
                    continue
                if not is_note(t):
                    continue
                # In the label columns, a bold cell is a heading and a row with
                # figures beside it is a label. Neither is a note.
                if c.column <= 2:
                    if (c.font and c.font.bold) or row_has_data(ws, c.row):
                        continue
                c.value = None
                cleared.append((ws.title, c.coordinate, t))
    wb.save(WB)

    for label, rows in (('cleared', cleared), ('shortened to "Check"', shortened),
                        ('title clause trimmed', titles)):
        print(f'--- {label}: {len(rows)}')
        for sheet, co, t in rows:
            print(f'    {sheet:<16} {co:<7} {t[:88]}')


if __name__ == '__main__':
    main()
