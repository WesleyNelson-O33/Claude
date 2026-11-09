"""Put the CTS palette into the deck's theme, and the text colours onto Charcoal.

The section dividers were never mine: slideLayout8 "CTS Section (navy)" has a
#122B82 background in her own template, and Deep Blue is on her list. What was
not CTS was the theme colour scheme underneath, which was still Office stock,
so the PowerPoint colour picker offered Office blue rather than any CTS colour.
"""
import re, shutil, zipfile, os

SRC = 'CTS Monthly Business Review - September 2026.pptx'
TMP = SRC + '.tmp'

SCHEME = """<a:clrScheme name="CTS"><a:dk1><a:srgbClr val="242C39"/></a:dk1>\
<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="122B82"/></a:dk2>\
<a:lt2><a:srgbClr val="E3E6EA"/></a:lt2><a:accent1><a:srgbClr val="005358"/></a:accent1>\
<a:accent2><a:srgbClr val="3FD0C9"/></a:accent2><a:accent3><a:srgbClr val="6D71FF"/></a:accent3>\
<a:accent4><a:srgbClr val="122B82"/></a:accent4><a:accent5><a:srgbClr val="F33844"/></a:accent5>\
<a:accent6><a:srgbClr val="8F4AFA"/></a:accent6><a:hlink><a:srgbClr val="6D71FF"/></a:hlink>\
<a:folHlink><a:srgbClr val="8F4AFA"/></a:folHlink></a:clrScheme>"""

# her palette has one text colour, Charcoal, so the two near blacks go to it
SWAP = {'1A1A1A': '242C39', '5F6670': '242C39'}

zin = zipfile.ZipFile(SRC)
zout = zipfile.ZipFile(TMP, 'w', zipfile.ZIP_DEFLATED)
n_theme = n_swap = 0
for item in zin.infolist():
    data = zin.read(item.filename)
    name = item.filename
    if name.startswith('ppt/theme/') and name.endswith('.xml'):
        x = data.decode('utf-8')
        new = re.sub(r'<a:clrScheme.*?</a:clrScheme>', SCHEME, x, count=1, flags=re.S)
        if new != x:
            n_theme += 1
        data = new.encode('utf-8')
    elif (name.startswith(('ppt/slides/slide', 'ppt/slideLayouts/', 'ppt/slideMasters/'))
          and name.endswith('.xml')):
        x = data.decode('utf-8')
        for old, newc in SWAP.items():
            c = x.count(f'val="{old}"')
            if c:
                x = x.replace(f'val="{old}"', f'val="{newc}"')
                n_swap += c
        data = x.encode('utf-8')
    zout.writestr(item, data)
zin.close(); zout.close()
shutil.move(TMP, SRC)
print(f'theme colour schemes rewritten: {n_theme}')
print(f'text colours moved to Charcoal: {n_swap}')
