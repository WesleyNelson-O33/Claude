"""Take the bullet dots off the commentary boxes.

The commentary paragraphs carried a bullet character on a hanging indent,
so the dot sat out in the left margin away from its own text. Danica asked
for them gone. This drops the bullet and closes the indent so the text sits
flush, which is how the rest of the deck reads.

Only the commentary paragraphs match. Every bullet character in the deck is
this one pattern, on the information slides. The cover, the section dividers
and the closing page have none.
"""
import re, shutil, sys, zipfile

PPR = re.compile(
    r'<a:pPr marL="342900" indent="-342900">'
    r'((?:(?!</a:pPr>).)*?)'
    r'<a:buSzPct val="100000"/><a:buChar char="•"/>'
    r'</a:pPr>',
    re.S)


def main(src, out):
    zin = zipfile.ZipFile(src)
    changed = {}
    total = 0
    for name in zin.namelist():
        if not re.match(r'ppt/slides/slide\d+\.xml$', name):
            continue
        xml = zin.read(name).decode('utf8')
        new, n = PPR.subn(r'<a:pPr marL="0" indent="0">\1<a:buNone/></a:pPr>', xml)
        if n:
            changed[name] = new.encode('utf8')
            total += n
            print('  %-18s %d bullets removed' % (name.split('/')[-1], n))

    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = changed.get(item.filename, zin.read(item.filename))
            zout.writestr(item, data)
    zin.close()
    print('\n%d bullets removed across %d slides' % (total, len(changed)))
    return total


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
