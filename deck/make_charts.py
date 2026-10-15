"""Draw the deck charts as HTML, in the CTS colours. Playwright shoots them."""
import json, os, html

D = os.path.dirname(os.path.abspath(__file__))
pl = json.load(open(f'{D}/data/pl.json'))
cust = json.load(open(f'{D}/data/customers.json'))
rec = json.load(open(f'{D}/data/receivables.json'))
mix = json.load(open(f'{D}/data/revmix.json'))

# The CTS palette, exactly as supplied.
TEAL, AQUA, GREY = '#005358', '#3FD0C9', '#E3E6EA'
MID, INK = '#242C39', '#242C39'          # Charcoal
BLUE, NAVY, RED, PURP = '#6D71FF', '#122B82', '#F33844', '#8F4AFA'
RULE = '#E3E6EA'

CSS = f"""
*{{margin:0;padding:0;box-sizing:border-box}}
body{{background:#fff;font-family:Calibri,'Segoe UI',Arial,sans-serif;color:{INK}}}
.wrap{{padding:14px 18px;display:inline-block}}
svg text{{font-family:Calibri,'Segoe UI',Arial,sans-serif}}
.lab{{font-size:13px;fill:{MID}}}
.val{{font-size:13px;fill:{INK};font-weight:700}}
.ax{{font-size:12px;fill:{MID}}}
.key{{font-size:13px;fill:{MID}}}
"""

def page(w, h, body):
    return (f'<!doctype html><html><head><meta charset="utf-8"><style>{CSS}</style></head>'
            f'<body><div class="wrap"><svg width="{w}" height="{h}" viewBox="0 0 {w} {h}">'
            f'{body}</svg></div></body></html>')

def money(x, dp=0):
    return f'{x:,.{dp}f}'

def legend(x, y, items):
    out, cx = [], x
    for col, lab in items:
        out.append(f'<rect x="{cx}" y="{y-9}" width="11" height="11" rx="2" fill="{col}"/>')
        out.append(f'<text class="key" x="{cx+16}" y="{y}">{html.escape(lab)}</text>')
        cx += 20 + len(lab) * 7.0
    return ''.join(out)

# ---------------------------------------------------- 1. revenue and EBITDA --
def chart_rev_ebitda():
    W, H = 1500, 250
    labs = pl['six']['labels']; rev = pl['six']['rev']; eb = pl['six']['ebitda']
    x0, y0, pw, ph = 70, 60, W - 110, H - 130
    top = max(max(rev), max(eb)) * 1.12
    bot = min(0, min(eb) * 1.35)
    span = top - bot
    sy = lambda v: y0 + ph - (v - bot) / span * ph
    n = len(labs); step = pw / n
    s = [f'<text x="0" y="24" class="val" font-size="15">Revenue and EBITDA, A$ thousands</text>']
    for k in range(5):
        v = bot + span * k / 4
        s.append(f'<line x1="{x0}" y1="{sy(v):.1f}" x2="{x0+pw}" y2="{sy(v):.1f}" stroke="{GREY}"/>')
        s.append(f'<text class="ax" x="{x0-10}" y="{sy(v)+4:.1f}" text-anchor="end">{v/1000:,.0f}</text>')
    for i, l in enumerate(labs):
        cx = x0 + step * i
        bw = step * 0.34
        s.append(f'<rect x="{cx+step*0.12:.1f}" y="{min(sy(rev[i]),sy(0)):.1f}" width="{bw:.1f}" '
                 f'height="{abs(sy(rev[i])-sy(0)):.1f}" fill="{TEAL}" rx="2"/>')
        s.append(f'<rect x="{cx+step*0.52:.1f}" y="{min(sy(eb[i]),sy(0)):.1f}" width="{bw:.1f}" '
                 f'height="{abs(sy(eb[i])-sy(0)):.1f}" fill="{AQUA}" rx="2"/>')
        s.append(f'<text class="lab" x="{cx+step/2:.1f}" y="{y0+ph+22}" text-anchor="middle">{l}</text>')
        s.append(f'<text class="val" font-size="11" x="{cx+step*0.12+bw/2:.1f}" '
                 f'y="{sy(rev[i])-6:.1f}" text-anchor="middle">{rev[i]/1000:,.0f}</text>')
        ey = sy(eb[i]) - 6 if eb[i] >= 0 else sy(eb[i]) + 14
        s.append(f'<text class="val" font-size="11" fill="{MID}" x="{cx+step*0.52+bw/2:.1f}" '
                 f'y="{ey:.1f}" text-anchor="middle">{eb[i]/1000:,.0f}</text>')
    s.append(f'<line x1="{x0}" y1="{sy(0):.1f}" x2="{x0+pw}" y2="{sy(0):.1f}" stroke="{MID}" stroke-width="1.4"/>')
    s.append(legend(x0, H - 22, [(TEAL, 'Revenue'), (AQUA, 'EBITDA')]))
    s.append(f'<text class="key" x="{x0+pw}" y="{H-22}" text-anchor="end">'
             f'Jul to Sep actual, Oct to Dec budget</text>')
    return page(W, H, ''.join(s))

# ------------------------------------------------- 2. by department ----------
def chart_dept():
    W, H = 1220, 700
    d = pl['dept']
    x0, y0, pw, ph = 130, 58, W - 180, H - 120
    vals = [x['inc_sep'] for x in d] + [x['res_sep'] for x in d]
    lo, hi = min(0, min(vals)) * 1.15, max(vals) * 1.12
    span = hi - lo
    sx = lambda v: x0 + (v - lo) / span * pw
    rowh = ph / len(d)
    s = [f'<text x="0" y="24" class="val" font-size="15">September revenue and result after overhead, A$ thousands</text>']
    s.append(f'<line x1="{sx(0):.1f}" y1="{y0}" x2="{sx(0):.1f}" y2="{y0+ph}" stroke="{MID}"/>')
    for i, x in enumerate(d):
        cy = y0 + rowh * i
        bh = rowh * 0.3
        s.append(f'<text class="lab" x="{x0-18}" y="{cy+rowh/2+4:.1f}" text-anchor="end">{x["name"]}</text>')
        for j, (v, col) in enumerate(((x['inc_sep'], TEAL), (x['res_sep'], AQUA))):
            yy = cy + rowh * (0.16 + j * 0.38)
            s.append(f'<rect x="{min(sx(v),sx(0)):.1f}" y="{yy:.1f}" width="{abs(sx(v)-sx(0)):.1f}" '
                     f'height="{bh:.1f}" fill="{col}" rx="2"/>')
            tx = sx(v) + (7 if v >= 0 else -7)
            anc = 'start' if v >= 0 else 'end'
            s.append(f'<text class="val" font-size="12" x="{tx:.1f}" y="{yy+bh*0.72:.1f}" '
                     f'text-anchor="{anc}">{v/1000:,.0f}</text>')
    s.append(legend(x0, H - 20, [(TEAL, 'Revenue'), (AQUA, 'Result after all overhead')]))
    return page(W, H, ''.join(s))

# ------------------------------------------------- 3. revenue mix ------------
def chart_mix():
    W, H = 1220, 260
    rows = [('September', mix['contract']['sep'], mix['noncontract']['sep'], mix['other']['sep'], mix['total']['sep']),
            ('Year to date', mix['contract']['ytd'], mix['noncontract']['ytd'], mix['other']['ytd'], mix['total']['ytd'])]
    x0, y0, pw = 130, 54, W - 180
    s = [f'<text x="0" y="24" class="val" font-size="15">Contracted against non-contract revenue</text>']
    for i, (lab, c, nc, ot, tot) in enumerate(rows):
        cy = y0 + i * 78
        s.append(f'<text class="lab" x="{x0-18}" y="{cy+28}" text-anchor="end">{lab}</text>')
        cx = x0
        for v, col, nm in ((c, TEAL, 'Contracted'), (nc, AQUA, 'Non-contract'), (ot, GREY, 'Other')):
            w = v / tot * pw
            s.append(f'<rect x="{cx:.1f}" y="{cy:.1f}" width="{w:.1f}" height="40" fill="{col}"/>')
            if w > 70:
                fc = '#fff' if col == TEAL else INK
                s.append(f'<text x="{cx+w/2:.1f}" y="{cy+25:.1f}" text-anchor="middle" '
                         f'font-size="13" font-weight="700" fill="{fc}">{v/tot*100:.0f}%</text>')
            cx += w
        s.append(f'<text class="key" x="{x0}" y="{cy+58}">{money(c/1000)}k contracted, '
                 f'{money(nc/1000)}k non-contract, {money(tot/1000)}k total</text>')
    s.append(legend(x0, H - 14, [(TEAL, 'Contracted'), (AQUA, 'Non-contract'), (GREY, 'Other income')]))
    return page(W, H, ''.join(s))

# ------------------------------------------------- 4. top 10 customers -------
def chart_top10():
    W, H = 1220, 620
    rows = sorted(cust['rows'], key=lambda r: -r['fy27'])
    tot = pl['pl']['Total Income']['ytd']
    x0, y0, pw, ph = 230, 58, W - 320, H - 110
    hi = max(r['fy27'] for r in rows) * 1.1
    rowh = ph / len(rows)
    s = [f'<text x="0" y="24" class="val" font-size="15">Top 10 customers, year to date revenue, A$ thousands</text>']
    for i, r in enumerate(rows):
        cy = y0 + rowh * i
        w = r['fy27'] / hi * pw
        s.append(f'<text class="lab" x="{x0-16}" y="{cy+rowh/2+5:.1f}" text-anchor="end">{html.escape(r["name"])}</text>')
        s.append(f'<rect x="{x0}" y="{cy+rowh*0.17:.1f}" width="{w:.1f}" height="{rowh*0.66:.1f}" fill="{TEAL}" rx="2"/>')
        s.append(f'<text class="val" font-size="13" x="{x0+w+9:.1f}" y="{cy+rowh/2+5:.1f}">'
                 f'{r["fy27"]/1000:,.0f}</text>')
        s.append(f'<text class="key" x="{W-20}" y="{cy+rowh/2+5:.1f}" text-anchor="end">'
                 f'{r["fy27"]/tot*100:.1f}%</text>')
    share = sum(r['fy27'] for r in rows) / tot * 100
    s.append(f'<text class="key" x="0" y="{H-14}">The ten together are {share:.0f} per cent of year to date revenue. '
             f'Share of total revenue shown at the right.</text>')
    return page(W, H, ''.join(s))

# ------------------------------------------------- 5. customers year on year -
def chart_yoy():
    W, H = 1220, 620
    rows = sorted(cust['rows'], key=lambda r: -r['fy27'])
    x0, y0, pw, ph = 230, 70, W - 330, H - 130
    hi = max(max(r['fy27'], r['fy26']) for r in rows) * 1.1
    rowh = ph / len(rows)
    s = [f'<text x="0" y="24" class="val" font-size="15">Top 10 customers, this year to date against last, A$ thousands</text>']
    for i, r in enumerate(rows):
        cy = y0 + rowh * i
        s.append(f'<text class="lab" x="{x0-16}" y="{cy+rowh/2+5:.1f}" text-anchor="end">{html.escape(r["name"])}</text>')
        for j, (v, col) in enumerate(((r['fy27'], TEAL), (r['fy26'], GREY))):
            w = v / hi * pw
            yy = cy + rowh * (0.12 + j * 0.42)
            s.append(f'<rect x="{x0}" y="{yy:.1f}" width="{max(w,1):.1f}" height="{rowh*0.36:.1f}" fill="{col}" rx="2"/>')
            s.append(f'<text class="val" font-size="11" fill="{INK if j else INK}" x="{x0+max(w,1)+7:.1f}" '
                     f'y="{yy+rowh*0.28:.1f}">{v/1000:,.0f}</text>')
        mv = r['fy27'] - r['fy26']
        col = TEAL if mv >= 0 else RED
        sign = '+' if mv >= 0 else ''
        s.append(f'<text font-size="12" font-weight="700" fill="{col}" x="{W-20}" '
                 f'y="{cy+rowh/2+5:.1f}" text-anchor="end">{sign}{mv/1000:,.0f}k</text>')
    s.append(legend(x0, H - 16, [(TEAL, cust['fy27_label']), (GREY, cust['fy26_label'])]))
    return page(W, H, ''.join(s))

# ------------------------------------------------- 6. aged receivables -------
def chart_ageing():
    W, H = 620, 430
    bk = [b for b in rec['buckets'] if b[1] > 0]
    x0, y0, pw, ph = 70, 56, W - 110, H - 120
    hi = max(b[1] for b in bk) * 1.15
    step = pw / len(bk)
    s = [f'<text x="0" y="24" class="val" font-size="15">Aged receivables, A$ thousands</text>']
    for i, (lab, v, pct) in enumerate(bk):
        cx = x0 + step * i
        h = v / hi * ph
        col = TEAL if lab == 'Current' else (AQUA if 'Under' in lab else RED)
        s.append(f'<rect x="{cx+step*0.2:.1f}" y="{y0+ph-h:.1f}" width="{step*0.6:.1f}" height="{h:.1f}" fill="{col}" rx="2"/>')
        s.append(f'<text class="val" font-size="12" x="{cx+step/2:.1f}" y="{y0+ph-h-8:.1f}" text-anchor="middle">{v/1000:,.0f}</text>')
        s.append(f'<text class="lab" font-size="11" x="{cx+step/2:.1f}" y="{y0+ph+20}" text-anchor="middle">{lab}</text>')
        s.append(f'<text class="key" font-size="11" x="{cx+step/2:.1f}" y="{y0+ph+36}" text-anchor="middle">{pct}%</text>')
    s.append(f'<line x1="{x0}" y1="{y0+ph}" x2="{x0+pw}" y2="{y0+ph}" stroke="{MID}"/>')
    s.append(f'<text class="key" x="0" y="{H-12}">{rec["invoice_count"]} invoices, {rec["customer_count"]} customers. '
             f'Overdue {rec["overdue_total"]/1000:,.0f}k, {rec["overdue_pct"]}% of the book.</text>')
    return page(W, H, ''.join(s))

# ------------------------------------------------- 7. utilisation ------------
def chart_util():
    W, H = 1220, 330
    u = [x for x in pl['util'] if x['ch'] + x['nc'] + x['lv'] > 0]
    x0, y0, pw, ph = 130, 56, W - 180, H - 110
    rowh = ph / len(u)
    s = [f'<text x="0" y="24" class="val" font-size="15">September hours by department</text>']
    for i, x in enumerate(u):
        cy = y0 + rowh * i
        tot = x['ch'] + x['nc'] + x['lv']
        s.append(f'<text class="lab" x="{x0-16}" y="{cy+rowh/2+5:.1f}" text-anchor="end">{x["name"]}</text>')
        cx = x0
        for v, col in ((x['ch'], TEAL), (x['nc'], AQUA), (x['lv'], GREY)):
            w = v / tot * pw
            s.append(f'<rect x="{cx:.1f}" y="{cy+rowh*0.2:.1f}" width="{w:.1f}" height="{rowh*0.6:.1f}" fill="{col}"/>')
            if w > 54:
                fc = '#fff' if col == TEAL else INK
                s.append(f'<text x="{cx+w/2:.1f}" y="{cy+rowh*0.58:.1f}" text-anchor="middle" '
                         f'font-size="12" font-weight="700" fill="{fc}">{v:,.0f}</text>')
            cx += w
        s.append(f'<text class="val" font-size="13" x="{W-20}" y="{cy+rowh/2+5:.1f}" text-anchor="end">'
                 f'{x["pct"]*100:.0f}% chargeable</text>')
    s.append(legend(x0, H - 14, [(TEAL, 'Chargeable'), (AQUA, 'Non-chargeable'), (GREY, 'Leave')]))
    return page(W, H, ''.join(s))

CHARTS = {'rev_ebitda': chart_rev_ebitda, 'dept': chart_dept, 'mix': chart_mix,
          'top10': chart_top10, 'yoy': chart_yoy, 'ageing': chart_ageing, 'util': chart_util}
for name, fn in CHARTS.items():
    open(f'{D}/html/{name}.html', 'w').write(fn())
    print('wrote', name)
