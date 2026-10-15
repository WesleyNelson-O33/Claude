"""Fill her September deck from the pack, Xero and the Playwright charts.

Nothing is invented. Every figure traces to the Financial Controller Pack
v18, to the Xero GL dump inside it, or to a Xero report named on the slide.
"""
import json, os, copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu

D = os.path.dirname(os.path.abspath(__file__))
pl = json.load(open(f'{D}/data/pl.json'))
sl = json.load(open(f'{D}/data/slides.json'))
cust = json.load(open(f'{D}/data/customers.json'))
rec = json.load(open(f'{D}/data/receivables.json'))
mix = json.load(open(f'{D}/data/revmix.json'))
lab = json.load(open(f'{D}/data/labour.json'))

prs = Presentation(f'{D}/deck_base.pptx')
S = {i: s for i, s in enumerate(prs.slides, 1)}

# ------------------------------------------------------------- helpers -----
def k(x, dp=0):
    """A$ thousands, brackets for negatives, the way the pack prints them."""
    v = x / 1000.0
    return f'({abs(v):,.{dp}f})' if v < -0.0005 else f'{v:,.{dp}f}'

def pct(x, dp=1):
    return f'{x*100:,.{dp}f}%'

def sgn(x, dp=0):
    v = x / 1000.0
    return f'{"+" if v >= 0 else ""}{v:,.{dp}f}'

def set_para(para, text):
    """Keep the first run's formatting, drop the rest."""
    runs = para.runs
    if not runs:
        para.text = text
        return
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)

def set_lines(shape, lines):
    """Rewrite a text box, reusing the paragraph formatting that is there."""
    tf = shape.text_frame
    paras = tf.paragraphs
    proto = paras[0]._p
    while len(tf.paragraphs) < len(lines):
        tf._txBody.append(copy.deepcopy(proto))
    for i, line in enumerate(lines):
        set_para(tf.paragraphs[i], line)
    for p in list(tf.paragraphs)[len(lines):]:
        p._p.getparent().remove(p._p)

from pptx.dml.color import RGBColor
TEAL_C, RED_C = RGBColor(0x00, 0x53, 0x58), RGBColor(0xA8, 0x05, 0x2E)

def colour_signs(table, cols):
    """The template coloured its own dummy variances. Recolour on our signs."""
    for r in range(1, len(table.rows)):
        for c in cols:
            if c >= len(table.columns):
                continue
            txt = table.cell(r, c).text.strip()
            if not txt or txt in ('new', 'no budget', 'n/a'):
                continue
            neg = txt.startswith('-') or txt.startswith('(')
            pos = txt.startswith('+')
            if not (neg or pos):
                continue
            for para in table.cell(r, c).text_frame.paragraphs:
                for run in para.runs:
                    run.font.color.rgb = RED_C if neg else TEAL_C

def fill(table, rows, start=1):
    for ri, row in enumerate(rows, start):
        for ci, val in enumerate(row):
            if val is None:
                continue
            cell = table.cell(ri, ci)
            set_para(cell.text_frame.paragraphs[0], str(val))
            for p in list(cell.text_frame.paragraphs)[1:]:
                p._p.getparent().remove(p._p)

def header(table, labels):
    for ci, val in enumerate(labels):
        if val is not None:
            set_para(table.cell(0, ci).text_frame.paragraphs[0], str(val))

def swap_chart(slide, img, shape=None, pad=0.0):
    """Drop the template's native chart and drop the Playwright shot in its place."""
    sh = shape or next(x for x in slide.shapes if x.has_chart)
    l, t, w, h = sh.left, sh.top, sh.width, sh.height
    sh._element.getparent().remove(sh._element)
    from PIL import Image
    iw, ih = Image.open(img).size
    box_ar, img_ar = w / h, iw / ih
    if img_ar > box_ar:                      # letterbox, never stretch
        nw, nh = w, int(w / img_ar)
        nl, nt = l, t + (h - nh) // 2
    else:
        nh, nw = h, int(h * img_ar)
        nl, nt = l + (w - nw) // 2, t
    return slide.shapes.add_picture(img, nl, nt, nw, nh)

def stat_cards(slide):
    """The three cards under the title: value box on top, label box under it.
    The eyebrow and the title sit at x=0.55, the cards start at 0.70, so the
    left edge is what tells them apart."""
    c = [x for x in slide.shapes if x.has_text_frame
         and Emu(x.left).inches >= 0.65 and Emu(x.left).inches < 6.8
         and 1.6 <= Emu(x.top).inches <= 2.45
         and x.text_frame.paragraphs[0].text.strip()]
    c.sort(key=lambda x: (round(Emu(x.top).inches, 2), Emu(x.left).inches))
    tops = sorted({round(Emu(x.top).inches, 2) for x in c})
    vals = [x for x in c if round(Emu(x.top).inches, 2) == tops[0]]
    labs = [x for x in c if round(Emu(x.top).inches, 2) != tops[0]]
    return vals, labs

def drop_rows(table, n):
    """Take n rows off the bottom so the table still fits its box."""
    trs = table._tbl.findall(
        '{http://schemas.openxmlformats.org/drawingml/2006/main}tr')
    for tr in trs[-n-1:-1]:          # keep the last row, it carries the total style
        table._tbl.remove(tr)

def tables(slide):
    return [x.table for x in slide.shapes if x.has_table]

def boxes(slide):
    return [x for x in slide.shapes if x.has_text_frame]

def commentary(slide, lines):
    """The right hand COMMENTARY panel is the last tall text box on the slide."""
    cand = [x for x in slide.shapes
            if x.has_text_frame and Emu(x.left).inches > 6.8 and Emu(x.height).inches > 2]
    set_lines(cand[-1], lines)

# =============================================== 4. P&L summary ============
s = S[4]
t = tables(s)[0]
by = {r['lab']: r for r in sl['s4']}
def ratio(num, den, fld):
    d = by[den][fld]
    return by[num][fld] / d if d else 0.0
rows = []
for r in sl['s4']:
    if r['pct']:
        num = 'Gross profit' if 'Gross' in r['lab'] else 'EBITDA'
        v = {f: ratio(num, 'Revenue', f) for f in ('sep', 'bsep', 'ytd', 'bytd', 'py')}
        rows.append([r['lab'], pct(v['sep']), pct(v['bsep']),
                     f"{(v['sep']-v['bsep'])*100:+.1f} pts", pct(v['ytd']), pct(v['bytd']),
                     f"{(v['ytd']-v['bytd'])*100:+.1f} pts", pct(v['py'])])
    else:
        rows.append([r['lab'], k(r['sep']), k(r['bsep']), sgn(r['sep'] - r['bsep']),
                     k(r['ytd']), k(r['bytd']), sgn(r['ytd'] - r['bytd']), k(r['py'])])
header(t, ['A$ thousands', 'Sep 26', 'Budget', 'Var', 'YTD', 'Budget', 'Var', 'Sep 25'])
fill(t, rows)
swap_chart(s, f'{D}/img/rev_ebitda.png')
colour_signs(t, [3, 6])
for b in boxes(s):
    if b.text_frame.paragraphs[0].text.startswith('Revenue and EBITDA'):
        set_para(b.text_frame.paragraphs[0],
                 'Revenue and EBITDA, Jul to Sep actual, Oct to Dec budget')
rev = sl['s4'][0]
commentary(s, [
    'Revenue beat budget by 148k. The year is still 372k behind.',
    'The whole beat is one Integration project, Aware Super. It does not repeat.',
    'Gross margin was 33.4 per cent, but that Integration work carries only 10.',
    'Overheads 230k against 232k budget. Cost control is not the problem.',
    'EBITDA negative 32k for the month, 6k YTD against a 223k budget.',
])

# =============================================== 5. per department =========
s = S[5]
t = tables(s)[0]
rows = [[d['name'], k(d['rev']), k(d['bud']), sgn(d['rev'] - d['bud']),
         k(d['gp']), pct(d['gm']), k(d['oh']), k(d['con'])] for d in sl['s5']]
tt = sl['s5_total']
rows.append(['CTS', k(tt['rev']), k(tt['bud']), sgn(tt['rev'] - tt['bud']),
             k(tt['gp']), pct(tt['gm']), k(tt['oh']), k(tt['con'])])
header(t, ['A$ thousands', 'Revenue', 'Budget', 'Var', 'Gross profit', 'GM %',
           'All overhead', 'Result'])
fill(t, rows[:5] + [['Admin', '-', '-', '-', '-', '', '-', '-']] + [rows[5]])
colour_signs(t, [3])
for b in boxes(s):
    if b.text_frame.paragraphs[0].text.startswith('Overheads are'):
        set_lines(b, ['All overhead is the department\'s own overhead plus its share of admin, '
                      'spread on the preset keys. Admin itself carries nothing once its '
                      'overhead is charged out. Result is revenue less direct costs less all '
                      'overhead, and the five departments add back to net profit exactly.'])
commentary(s, [
    'Support is the only department covering its overhead. It made 29k.',
    'Events lost 54k. It earned 73k and was charged 111k of overhead.',
    'Integration made 10k on 235k. A 10 per cent margin is thin.',
    'Consulting lost 20k, carrying overhead with almost no fee base.',
    'The five results add back to net profit exactly.',
])

# =============================================== 6. by department chart ====
swap_chart(S[6], f'{D}/img/dept.png')
commentary(S[6], [
    'Support and Integration were 85 per cent of September revenue.',
    "Integration's bar is almost all Aware Super. Strip it out and it changes.",
    'Only those two clear their overhead charge. The other three do not.',
    'Events carries the biggest charge because of the Office Dept key.',
    'That key is fixed for the year. Revisit it at FY28, not now.',
])

# =============================================== 7. risk revenue ===========
s = S[7]
con_pct = mix['contract']['ytd'] / mix['total']['ytd']
non_pct = mix['noncontract']['ytd'] / mix['total']['ytd']
flagged = sum(1 for r in sl['s7'] if r['risk'] in ('High', 'Medium'))
stats = [(f'{con_pct*100:.0f}%', 'Contracted revenue, YTD'),
         (f'{non_pct*100:.0f}%', 'Non-contract revenue, at risk'),
         (str(flagged), 'Revenue lines flagged Medium or High')]
vals, labs = stat_cards(s)
for i, (v, l) in enumerate(stats):
    if i < len(vals): set_para(vals[i].text_frame.paragraphs[0], v)
    if i < len(labs): set_para(labs[i].text_frame.paragraphs[0], l)
swap_chart(s, f'{D}/img/mix.png')
t = tables(s)[0]
header(t, ['Revenue line, YTD', 'Actual', 'Budget', 'Var', 'Var %', 'Risk'])
fill(t, [[r['lab'], k(r['act']), k(r['bud']), sgn(r['var']),
          (pct(r['pct']) if r['bud'] else 'no budget'), r['risk']] for r in sl['s7']])
colour_signs(t, [3, 4])
commentary(s, [
    '55 per cent of revenue is non-contract and must be won again.',
    'Integration non-contract is 254k behind budget for the year.',
    'Consulting non-contract is 233k behind. That budget needs reforecasting.',
    'Support contract income is 134k behind. Scope, or a renewal we lost.',
    'Two lines earn with no budget mapped. A pack fix, not trading.',
])

# =============================================== 9. top 10 table ===========
s = S[9]
t = tables(s)[0]
tot = pl['pl']['Total Income']['ytd']
rows_c = sorted(cust['rows'], key=lambda r: -r['fy27'])
rr = []
for i, c in enumerate(rows_c, 1):
    mv = c['fy27'] - c['fy26']
    yo = 'new' if c['fy26'] == 0 else f'{mv/c["fy26"]*100:+.0f}%'
    rr.append([f'{i}.  {c["name"]}', k(c['fy27']), f'{c["fy27"]/tot*100:.0f}%',
               k(c['fy26']), yo, sgn(mv)])
tsum = sum(c['fy27'] for c in rows_c); psum = sum(c['fy26'] for c in rows_c)
rr.append(['Top 10', k(tsum), f'{tsum/tot*100:.0f}%', k(psum),
           f'{(tsum-psum)/psum*100:+.0f}%', sgn(tsum - psum)])
header(t, ['Customer, A$ thousands', 'FY27 YTD', 'Share of YTD', 'FY26 YTD', 'YoY %', 'YoY A$k'])
fill(t, rr)
colour_signs(t, [4, 5])
for b in boxes(s):
    if b.text_frame.paragraphs[0].text.startswith('Top 10 customers are'):
        set_lines(b, [f'FY27 is earned revenue from the Xero general ledger, ex GST, and ties to '
                      f'the P&L. FY26 is invoices issued 1 Jul to 30 Sep 2025 from Xero, ex GST.'])
top3 = sum(c['fy27'] for c in rows_c[:3])
commentary(s, [
    'The top ten are 76 per cent of revenue. The top three, 49.',
    'Aware Super went from nil to 285k. All of it one fit-out.',
    'Deloitte is the only large recurring customer, steady at 402k.',
    'Three of the ten are new, worth 368k between them.',
    'Lose Deloitte and a quarter of the book goes with it.',
])

# =============================================== 10. top 10 chart ==========
s = S[10]
for b in boxes(s):
    p0 = b.text_frame.paragraphs[0].text
    if p0.startswith('Top customers, FY27'):
        set_para(b.text_frame.paragraphs[0], 'Top 10 customers by revenue')
    if p0.startswith('A$ thousands, same period'):
        set_lines(b, ['Earned revenue from the Xero general ledger, ex GST, 1 Jul to 30 Sep 2026. '
                      'Ties to total income on the P&L.'])
swap_chart(s, f'{D}/img/top10.png')
commentary(s, [
    'Deloitte and Aware Super alone are 42 per cent of the year.',
    'Below the top five the book falls away fast.',
    'Aware Super is one project. Treat it as one-off when forecasting.',
    'No customer in the ten could replace Deloitte.',
    'The other eighteen customers are worth 384k between them.',
])

# =============================================== 11. receivables ===========
s = S[11]
dso = rec['total_outstanding'] / (pl['pl']['Total Income']['ytd'] / 92)
over60 = sum(b[1] for b in rec['buckets'] if b[0] in ('2 months', '3 months', 'Over 3 months'))
stats = [(f'{dso:.0f} days', 'DSO, target 30 days'),
         (f'${rec["total_outstanding"]/1e6:.2f}m', 'Trade receivables at month end'),
         (f'${over60/1000:.0f}k', 'Over 60 days')]
vals, labs = stat_cards(s)
for i, (v, l) in enumerate(stats):
    if i < len(vals): set_para(vals[i].text_frame.paragraphs[0], v)
    if i < len(labs): set_para(labs[i].text_frame.paragraphs[0], l)
for b in boxes(s):
    p0 = b.text_frame.paragraphs[0].text
    if p0.startswith('DSO, last six months'):
        set_para(b.text_frame.paragraphs[0], 'Aged receivables at 30 September')
    if p0.startswith('Aged receivables, A$'):
        set_para(b.text_frame.paragraphs[0], 'Who owes the most, A$ thousands')
    if p0.startswith('DSO = receivables'):
        set_lines(b, ['DSO = receivables / (last three months revenue / 92 days). '
                      'Source: Xero aged receivables at 30 September 2026.'])
swap_chart(s, f'{D}/img/ageing.png')
t = tables(s)[0]
header(t, ['Customer', 'A$k', 'Share'])
fill(t, [[n, k(v), f'{p:.1f}%'] for n, v, p in rec['top_debtors']] +
        [['All others', k(rec['total_outstanding'] - sum(d[1] for d in rec['top_debtors'])),
          f'{100-sum(d[2] for d in rec["top_debtors"]):.1f}%']])
commentary(s, [
    'Receivables are 833k across 61 invoices and 28 customers.',
    'DSO is 47 days against a 30 day target.',
    '29 per cent of the book is overdue, 245k of it.',
    'Aware Super is a third of the ledger, on one invoice of 226k.',
    'Two invoices sit over 90 days. Property NSW 20k, Meet Magic 5k.',
])

# =============================================== 12. customers YoY =========
s = S[12]
for b in boxes(s):
    if b.text_frame.paragraphs[0].text.startswith('A$ thousands, same period'):
        set_lines(b, ['FY27 is earned revenue from the Xero general ledger, ex GST. FY26 is '
                      'invoices issued in the same period from Xero, ex GST. Ranked by this year.'])
swap_chart(s, f'{D}/img/yoy.png')
commentary(s, [
    'Revenue is up on last year, but the growth is three new customers.',
    'Aware Super, Data#3 and Generation-e are 368k of new revenue.',
    'On the customers we had last year the picture is flat.',
    'ASX is down a third, from 90k to 61k, with no replacement.',
    'If the three do not repeat, FY28 starts 368k behind.',
])

# =============================================== 14. utilisation ===========
s = S[14]
t = tables(s)[0]
u = {x['name']: x for x in pl['util']}
REPORTING = [n for n in ['Support', 'Events', 'Video', 'Integration', 'Consulting']
             if u[n]['ch'] + u[n]['nc'] + u[n]['lv'] > 0]
drop_rows(t, len(pl['util']) - len(REPORTING))
rows = [[n, f"{u[n]['fte']:.1f}", f"{u[n]['ch']:,.0f}", f"{u[n]['nc']:,.0f}",
         f"{u[n]['lv']:,.0f}", pct(u[n]['pct']), '', ''] for n in REPORTING]
tc = sum(x['ch'] for x in pl['util']); tn = sum(x['nc'] for x in pl['util'])
tl = sum(x['lv'] for x in pl['util']); tf = sum(x['fte'] for x in pl['util'])
rows.append(['CTS', f'{tf:.1f}', f'{tc:,.0f}', f'{tn:,.0f}', f'{tl:,.0f}',
             pct(tc / (tc + tn) if tc + tn else 0), '', ''])
header(t, ['September', 'FTE', 'Chargeable', 'Non-chargeable', 'Leave', 'Chg %',
           'Prior', 'vs target'])
fill(t, rows)
swap_chart(s, f'{D}/img/util.png')
commentary(s, [
    'Support shows 100 per cent because no non-chargeable hours are recorded.',
    'That is a reporting gap. Fix it before the next board pack.',
    'Events is 48 per cent, with more leave than chargeable hours.',
    'Consulting is 33 per cent on 40 chargeable hours.',
    'Video and Integration do not report hours at all.',
])

# =============================================== 15. profit per hour =======
s = S[15]
t = tables(s)[0]
dep = {d['name']: d for d in sl['s5']}
drop_rows(t, len(pl['util']) - len(REPORTING))
rows = []
for n in REPORTING:
    x = u[n]
    rev = dep[n]['rev']; lc = lab[n]['lab_sep']
    rph = rev / x['ch']
    five = x['ch'] / x['pct'] * 0.05 * rph if x['pct'] else 0
    rows.append([n, pct(x['pct']), f"{x['ch']:,.0f}", f'${rph:,.0f}', k(rev), k(lc),
                 k(rev + lc), f'${(rev+lc)/x["ch"]:,.0f}', sgn(five)])
rows.append(['CTS', pct(tc / (tc + tn)), f'{tc:,.0f}', '',
             k(sum(dep[n]['rev'] for n in dep)), k(sum(lab[n]['lab_sep'] for n in lab)),
             k(sum(dep[n]['rev'] for n in dep) + sum(lab[n]['lab_sep'] for n in lab)), '', ''])
header(t, ['September', 'Chg %', 'Chg hrs', 'Rev / hr', 'Revenue',
           'Labour', 'After', 'Per hr', '+5 pts'])
fill(t, rows)
for b in boxes(s):
    p0 = b.text_frame.paragraphs[0].text
    if p0.startswith('+$55k'):
        set_para(b.text_frame.paragraphs[0], '+$13k a month')
    if p0.startswith('Five more points'):
        set_lines(b, ['Five more points of utilisation across the three departments that report '
                      'hours, at the revenue per chargeable hour each already earns'])
    if p0.startswith('Labour revenue ='):
        set_lines(b, ['Video and Integration are left out: they do not report hours. Revenue per '
                      'hour is department revenue divided by chargeable hours. It is not a charge '
                      'rate, it includes equipment, hire and freight, and the pack does not hold '
                      'charge rates. Labour is wages, salaries, superannuation and workers '
                      'compensation coded to the department.'])
commentary(s, [
    'Support earns $114 a chargeable hour and keeps $54 after labour.',
    'Events shows $276 an hour, but that is equipment, not a rate.',
    'Consulting is $88 an hour on only 40 hours. Too thin to read.',
    'Five points of utilisation is worth about 13k a month.',
    'This cannot be done properly until every department reports hours.',
])

# =============================================== 17. actions ===============
s = S[17]
t = tables(s)[0]
header(t, ['Area', 'What we agreed', 'Owner', 'Due'])
fill(t, [
    ['P&L', 'September beat budget on one Integration project. Reforecast the second half on '
            'the recurring base, not on this month.', 'GC and FC', '31 Oct 26'],
    ['Risk revenue', 'Integration and Consulting non-contract budgets are 487k behind for the '
                     'year. Agree a reforecast or a plan to close it.', 'Sales and GC', '31 Oct 26'],
    ['Customers and DSO', 'DSO is 47 days against a 30 day target. Chase Property NSW 19.7k and '
                          'Meet Magic 5.1k, both over 90 days.', 'FC', '24 Oct 26'],
    ['Utilisation', 'Support is recording no non-chargeable hours and Video and Integration '
                    'report none at all. Fix the reporting before the next review.', 'Ops', '14 Nov 26'],
])

out = f'{D}/CTS Monthly Business Review - September 2026.pptx'
prs.save(out)
print('saved', out)
