"""Make every data bar carry Danica's own definition, in both programs.

A data bar rule has two halves: the old 2007 half, and a newer x14 half that
holds the axis and the negative colour. openpyxl writes the 2007 half and an
id pointing at the x14 half, but never writes that half. Each program then
invents one. LibreOffice fills in no axis and green negatives; Excel falls
back to 2007 behaviour, one colour drawn from the left. The controller pack
and the deck ended up showing the same rule two different ways.

This writes the x14 half, copied from Danica's rule on the Jul-Sept sheet of
her utilisation report: bar and border in the rule's own colour, negatives in
red, a black axis placed automatically, bars the full width of the cell. Any
definition another program has already invented is replaced, so the two agree.

Run it after tools/recalc.py, and after anything else that saves the workbook
through openpyxl, before taking a snip.
"""
import os, re, shutil, sys, tempfile, zipfile

SHEET = re.compile(r'xl/worksheets/sheet\d+\.xml$')
X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
XM = 'http://schemas.microsoft.com/office/excel/2006/main'
CF_EXT = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}'
BLOCK = re.compile(r'<conditionalFormatting[^>]*sqref="([^"]+)"[^>]*>(.*?)</conditionalFormatting>', re.S)
RULE = re.compile(r'<cfRule[^>]*type="dataBar".*?</cfRule>', re.S)
ID = re.compile(r'<x14:id>(\{[0-9A-Fa-f-]+\})</x14:id>')
COLOR = re.compile(r'<color rgb="([0-9A-Fa-f]{6,8})"\s*/?>')
EXISTING = re.compile(r'<x14:conditionalFormatting[ >].*?</x14:conditionalFormatting>', re.S)
CONTAINER = re.compile(r'<x14:conditionalFormattings>(.*?)</x14:conditionalFormattings>', re.S)


def definition(guid, colour, sqref):
    """Her settings, verbatim apart from the colour and the range."""
    return (
        '<x14:conditionalFormatting xmlns:xm="%s">'
        '<x14:cfRule type="dataBar" id="%s">'
        '<x14:dataBar minLength="0" maxLength="100" border="1" '
        'negativeBarBorderColorSameAsPositive="0">'
        '<x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>'
        '<x14:borderColor rgb="%s"/>'
        '<x14:negativeFillColor rgb="FFFF0000"/>'
        '<x14:negativeBorderColor rgb="FFFF0000"/>'
        '<x14:axisColor rgb="FF000000"/>'
        '</x14:dataBar></x14:cfRule>'
        '<xm:sqref>%s</xm:sqref>'
        '</x14:conditionalFormatting>' % (XM, guid, colour, sqref))


def fix_sheet(x):
    wanted = []
    for sqref, body in BLOCK.findall(x):
        for rule in RULE.findall(body):
            g, c = ID.search(rule), COLOR.search(rule)
            if g and c:
                wanted.append((g.group(1), c.group(1), sqref))
    if not wanted:
        return x, []
    ids = {g for g, _, _ in wanted}
    mine = ''.join(definition(*w) for w in wanted)
    if ids and all(definition(*w) in x for w in wanted):
        return x, []                       # already exactly this

    if '<x14:conditionalFormattings>' in x:
        # drop whatever is defined for these ids, keep anything else
        def trim(m):
            kept = [b for b in EXISTING.findall(m.group(1))
                    if not any(('id="%s"' % i) in b for i in ids)]
            return '<x14:conditionalFormattings>%s%s</x14:conditionalFormattings>' % (
                ''.join(kept), mine)
        x = CONTAINER.sub(trim, x, count=1)
    else:
        ext = ('<ext xmlns:x14="%s" uri="%s"><x14:conditionalFormattings>%s'
               '</x14:conditionalFormattings></ext>' % (X14, CF_EXT, mine))
        if '</extLst>' in x:
            i = x.rindex('</extLst>')
            x = x[:i] + ext + x[i:]
        else:
            i = x.rindex('</worksheet>')
            x = x[:i] + '<extLst>' + ext + '</extLst>' + x[i:]
    return x, wanted


def fix(path, quiet=False):
    path = os.path.abspath(path)
    with zipfile.ZipFile(path) as z:
        items = [(i, z.read(i.filename)) for i in z.infolist()]
    done, out = [], []
    for info, data in items:
        if SHEET.search(info.filename):
            new, got = fix_sheet(data.decode('utf-8'))
            if got:
                data = new.encode('utf-8')
                done += [(info.filename, g[2], g[1]) for g in got]
        out.append((info, data))
    if not done:
        if not quiet:
            print('data bars already match in %s' % os.path.basename(path))
        return 0
    tmp = tempfile.mktemp(suffix='.xlsx')
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        for info, data in out:
            z.writestr(info, data)
    shutil.move(tmp, path)
    if not quiet:
        for part, sqref, colour in done:
            print('  %s %s: data bar set to her definition, colour %s'
                  % (part, sqref, colour))
    return len(done)


if __name__ == '__main__':
    for p in sys.argv[1:]:
        fix(p)
