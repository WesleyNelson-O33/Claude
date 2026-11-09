"""Three colour schemes for the P&L, drawn as the sheet will actually look."""
import os, json
D = os.path.dirname(os.path.abspath(__file__))
B = json.load(open(f'{D}/../pack/data/brand.json'))
h = lambda k: '#' + B[k]
TEAL, AQUA, NAVY, BLUE = h('dark_teal'), h('bright_teal'), h('deep_blue'), h('primary_blue')
INK, RED, GREY, GREYL, W = h('charcoal'), h('accent_red'), h('grey'), h('grey_lt'), '#FFFFFF'

# label, kind, sep, aug, ytd      kind: head|band|sub|leaf|sect|res|pct
ROWS = [
 ('Line',                        'head', 'Sep 26', 'Aug 26', 'YTD'),
 ('Total Income',                'band', '561,945', '531,733', '1,636,138'),
 ('    Production',              'sub',  '80,607',  '112,142', '268,683'),
 ('        Events',              'leaf2','73,304',  '104,930', '248,951'),
 ('            Contract Income', 'leaf', '21,901',  '19,318',  '62,406'),
 ('            Non Contract Income','leaf','47,238','81,142',  '175,124'),
 ('    Cost of Sales',           'sect', '374,130', '349,167', '1,079,146'),
 ('    Production',              'sub',  '18,529',  '43,106',  '98,221'),
 ('        Events',              'leaf2','16,423',  '41,002',  '92,340'),
 ('            Sub-Contract Labour','leaf','9,117',  '22,845',  '48,990'),
 ('            Equipment - Hires','leaf', '(412)',   '3,910',   '9,845'),
 ('    Gross profit',            'res',  '187,814', '182,566', '556,992'),
 ('    Gross profit %',          'pct',  '33.4%',   '34.3%',   '34.0%'),
 ('    EBITDA',                  'res',  '(32,437)','(18,904)','(61,215)'),
 ('    EBITDA %',                'pct',  '(5.8%)',  '(3.6%)',  '(3.7%)'),
 ('    Net profit',              'res',  '(41,464)','(27,118)','(88,330)'),
 ('    Net profit %',            'pct',  '(7.4%)',  '(5.1%)',  '(5.4%)'),
]

def neg(v):
    return v.startswith('(')

OPTS = [
 ('Option 1  —  Charcoal', 'The chrome is charcoal, which is the one colour in your '
  'palette that is not a department. Bright teal only as a hairline.',
  dict(title=INK, head_bg=INK, head_fg=W, rule=AQUA,
       band_bg=GREY, band_fg=INK, sub_bg=GREYL, sub_fg=INK,
       sect_bg=GREY, sect_fg=INK, res_bg=GREYL, res_fg=INK)),
 ('Option 2  —  Deep Blue', 'The chrome is deep blue. The most conventional look for '
  'a P&L, and clearly not the logo green.',
  dict(title=NAVY, head_bg=NAVY, head_fg=W, rule=BLUE,
       band_bg=GREY, band_fg=NAVY, sub_bg=GREYL, sub_fg=NAVY,
       sect_bg=GREY, sect_fg=NAVY, res_bg=GREYL, res_fg=NAVY)),
 ('Option 3  —  Charcoal with a deep blue band', 'Charcoal heading, and the top line '
  'and the result lines picked out in deep blue so the eye lands on them.',
  dict(title=INK, head_bg=INK, head_fg=W, rule=NAVY,
       band_bg=NAVY, band_fg=W, sub_bg=GREYL, sub_fg=INK,
       sect_bg=GREY, sect_fg=INK, res_bg='#DDE1F0', res_fg=NAVY)),
]

CSS = """
*{margin:0;padding:0;box-sizing:border-box}
body{background:#fff;font-family:Calibri,'Segoe UI',Arial,sans-serif;color:%s}
.wrap{padding:20px 22px;display:inline-block}
h2{font-size:17px;margin:0 0 3px}
.sub{font-size:12.5px;color:#5b6472;margin-bottom:10px;max-width:470px}
table{border-collapse:collapse;width:470px;margin-bottom:26px}
td{font-size:12.5px;padding:3px 9px;border-bottom:1px solid %s;white-space:nowrap}
td.n{text-align:right;font-variant-numeric:tabular-nums}
.lbl{font-size:11px;letter-spacing:.09em;text-transform:uppercase;color:#5b6472;margin:0 0 7px}
""" % (INK, GREY)

def table(o):
    out = []
    for lab, kind, a, b, c in ROWS:
        st = bg = ''
        fg = INK
        bold = False
        if kind == 'head':
            bg, fg, bold = o['head_bg'], o['head_fg'], True
            st = f";border-bottom:3px solid {o['rule']}"
        elif kind == 'band':
            bg, fg, bold = o['band_bg'], o['band_fg'], True
        elif kind == 'sect':
            bg, fg, bold = o['sect_bg'], o['sect_fg'], True
        elif kind == 'sub':
            bg, fg, bold = o['sub_bg'], o['sub_fg'], True
        elif kind in ('res', 'pct'):
            bg, fg, bold = o['res_bg'], o['res_fg'], True
        elif kind == 'leaf2':
            bold = True
        style = f'background:{bg};color:{fg};font-weight:{700 if bold else 400}{st}'
        pad = lab.count(' ') and (len(lab) - len(lab.lstrip())) * 4
        cells = [f'<td style="{style};padding-left:{9 + (pad or 0)}px">{lab.strip()}</td>']
        for v in (a, b, c):
            # numbers are charcoal, except negatives, which are red; on a dark
            # band the text is white and a negative is the pale red
            if kind == 'head':
                col = fg
            elif neg(v):
                col = '#FF9DA6' if bg in (o['head_bg'], NAVY) and fg == W else RED
            else:
                col = fg
            cells.append(f'<td class="n" style="{style};color:{col}">{v}</td>')
        out.append('<tr>' + ''.join(cells) + '</tr>')
    return '<table>' + ''.join(out) + '</table>'

body = []
for name, note, o in OPTS:
    body.append(f'<h2 style="color:{o["title"]}">{name}</h2><p class="sub">{note}</p>'
                + table(o))
open(f'{D}/html/plopts.html', 'w').write(
    f'<!doctype html><html><head><meta charset="utf-8"><style>{CSS}</style></head>'
    f'<body><div class="wrap">{"".join(body)}</div></body></html>')
print('wrote plopts')
