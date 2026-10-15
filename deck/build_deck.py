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

def squeeze(shape, table):
    """Hold the table inside its box. LibreOffice and PowerPoint both grow a
    table to fit its text, so the row heights have to be set, not assumed."""
    rows = table.rows
    h = int(shape.height / len(rows))
    for r in rows:
        r.height = h

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
    'Break even needs 549k of revenue a month. We are running at 541k.',
    'The EBITDA miss is all margin. Overheads came in 48k under budget.',
    'Gross profit missed 280k: 170k on volume, 110k on a 6.8 point margin drop.',
    'At this run rate we meet the revenue budget and still lose 37k.',
    'Reforecast on margin, not revenue. Cost is not the lever here.',
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
    'Only Support covers its overhead. It returned 12 per cent on revenue.',
    'Events is charged 87.5 per cent of premises on 13 per cent of revenue.',
    'On that key Events cannot be profitable, whatever it sells.',
    'Integration made 10k on 235k. A fit-out at that margin carries risk for nothing.',
    'Consulting holds 20k of overhead on a 3.5k fee base. Its future needs a decision.',
])

# =============================================== 6. by department chart ====
swap_chart(S[6], f'{D}/img/dept.png')
commentary(S[6], [
    'Two of five departments clear their overhead. Three do not.',
    "Integration's bar is one project. Underlying it is a 60k a quarter unit.",
    'Events is the structural problem: heavy overhead, thin and falling fee income.',
    'The overhead keys were set for a different shape of business.',
    'Test the Office Dept key before the FY28 budget, not mid year.',
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
    'Only 43 per cent of revenue is contracted. The rest is rebuilt every month.',
    'Integration and Consulting non-contract are 487k behind for the quarter.',
    'That budget is not recoverable across the remaining nine months.',
    'Recommend a formal reforecast of both lines before the half year.',
    'Two lines earn against no budget. A pack fix, not a trading issue.',
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
    'Top three customers are 49 per cent of revenue. That is the real exposure.',
    'Deloitte at 402k is close to the whole EBITDA base. Losing it ends the year.',
    'Aware Super at 285k is one fit-out and will not be there next quarter.',
    'Three of the ten are new, 368k, none contracted beyond the current job.',
    'Recommend a retention plan on Deloitte and contract cover on the new three.',
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
    'The concentration is in project work, not in contracted revenue.',
    'Deloitte and Aware Super are 42 per cent of the year between them.',
    'Below the top five the book falls away fast. Number ten is 37k.',
    'Everything outside the top ten is 384k, with no depth behind it.',
    'Put the business development target on the second tier, not the top.',
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
    'DSO is 47.5 days against a 30 day target.',
    'Every day of DSO is 17.7k of cash. Seventeen days is 309k.',
    '309k sits in debtors against net debt of 723k. Fastest cash lever we have.',
    'Aware Super is 277k of the 833k, on one invoice due 9 October.',
    'Two invoices over 90 days, 25k. Chase or provide this month.',
])

# =============================================== 12. customers YoY =========
s = S[12]
for b in boxes(s):
    if b.text_frame.paragraphs[0].text.startswith('A$ thousands, same period'):
        set_lines(b, ['FY27 is earned revenue from the Xero general ledger, ex GST. FY26 is '
                      'invoices issued in the same period from Xero, ex GST. Ranked by this year.'])
swap_chart(s, f'{D}/img/yoy.png')
commentary(s, [
    'All the growth is three new customers. The base we had is flat.',
    'Aware Super, Data#3 and Generation-e are 368k and none existed last year.',
    'On the existing base, gains of 146k against losses of 50k.',
    'ASX is down a third. That is a relationship to repair, not a market shift.',
    'If the three new names do not repeat, FY28 opens 368k behind.',
])

# =============================================== 14. utilisation ===========
s = s_ = S[14]
t = tables(s)[0]
u = {x['name']: x for x in pl['util']}
REPORTING = [n for n in ['Support', 'Events', 'Video', 'Integration', 'Consulting']
             if u[n]['ch'] + u[n]['nc'] + u[n]['lv'] > 0]
rows = [[n, f"{u[n]['fte']:.1f}", f"{u[n]['ch']:,.0f}", f"{u[n]['nc']:,.0f}",
         f"{u[n]['lv']:,.0f}", pct(u[n]['pct']), '', ''] for n in REPORTING]
tc = sum(x['ch'] for x in pl['util'])
tn = sum(x['nc'] for x in pl['util']) + pl['util_admin']['nc']
tl = sum(x['lv'] for x in pl['util']); tf = sum(x['fte'] for x in pl['util'])
rows.append(['CTS', f'{tf:.1f}', f'{tc:,.0f}', f'{tn:,.0f}', f'{tl:,.0f}',
             pct(tc / (tc + tn) if tc + tn else 0), '', ''])
header(t, ['September', 'FTE', 'Chargeable', 'Non-chargeable', 'Leave', 'Chg %',
           'Prior', 'vs target'])
squeeze([x for x in s_.shapes if x.has_table][0], t)
fill(t, rows)
_pic = swap_chart(s, f'{D}/img/util.png')
_pic.top, _pic.height = Inches(4.02), Inches(1.00)
_pic.left, _pic.width = Inches(0.55), Inches(6.10)
commentary(s, [
    'Support records 11 non-chargeable hours in a month. That is not credible.',
    'Until it is fixed the 99.5 per cent means nothing, and nor does the target.',
    'Events is at 49 per cent with 333 hours of leave. Capacity is sitting idle.',
    'Consulting at 27 per cent confirms there is no fee base to utilise.',
    'Hours are split by the job, the same basis revenue is split on.',
])

# =============================================== 15. profit per hour =======
s = s_ = S[15]
t = tables(s)[0]
dep = {d['name']: d for d in sl['s5']}
THICK = [n for n in REPORTING if u[n]['ch'] >= 100]
drop_rows(t, len(REPORTING) - len(THICK))
rows = []
for n in THICK:
    x = u[n]
    rev = dep[n]['rev']; lc = lab[n]['lab_sep']
    rph = rev / x['ch']
    five = x['ch'] / x['pct'] * 0.05 * rph if x['pct'] else 0
    rows.append([n, pct(x['pct']), f"{x['ch']:,.0f}", f'${rph:,.0f}', k(rev), k(lc),
                 k(rev + lc), f'${(rev+lc)/x["ch"]:,.0f}', sgn(five)])
rows.append(['CTS', pct(tc / (tc + tn)), f'{tc:,.0f}', '',
             k(sum(dep[n]['rev'] for n in dep)), k(sum(lab[n]['lab_sep'] for n in lab)),
             k(sum(dep[n]['rev'] for n in dep) + sum(lab[n]['lab_sep'] for n in lab)), '', ''])
_five = sum(u[n]['ch'] / u[n]['pct'] * 0.05 * (dep[n]['rev'] / u[n]['ch'])
            for n in THICK if u[n]['pct'])
header(t, ['September', 'Chg %', 'Chg hrs', 'Rev / hr', 'Revenue',
           'Labour', 'After', 'Per hr', '+5 pts'])
squeeze([x for x in s_.shapes if x.has_table][0], t)
fill(t, rows)
for b in boxes(s):
    p0 = b.text_frame.paragraphs[0].text
    if p0.startswith('+$55k') or p0.startswith('+$13k'):
        set_para(b.text_frame.paragraphs[0], f'+${_five/1000:,.0f}k a month')
    if p0.startswith('Five more points'):
        set_lines(b, ['Five more points of utilisation on Support and Events, at the revenue '
                      'per chargeable hour each already earns, before any extra cost'])
    if p0.startswith('Labour revenue ='):
        set_lines(b, ['Video and Integration are left out: 24 and 10 chargeable hours are too '
                      'few for a rate to mean anything. Revenue per '
                      'hour is department revenue divided by chargeable hours. It is not a charge '
                      'rate, it includes equipment, hire and freight, and the pack does not hold '
                      'charge rates. Labour is wages, salaries, superannuation and workers '
                      'compensation coded to the department.'])
commentary(s, [
    'Support returns $54 of margin an hour after direct labour. That is the engine.',
    'Events returns $258, but on equipment revenue. It is not a labour rate.',
    'Integration billed 235k on 9.5 chargeable hours. That work was bought in.',
    'That is why its margin is 10 per cent. We are reselling, not delivering.',
    'Five points on Support and Events is worth 20k a month at no extra cost.',
])

# =============================================== 17. actions ===============
s = S[17]
t = tables(s)[0]
header(t, ['Area', 'What we agreed', 'Owner', 'Due'])
fill(t, [
    ['P&L', 'Reforecast FY27 on margin, not revenue. At the current run rate we meet the '
            'revenue budget and deliver an EBITDA loss of 37k against a 397k plan.',
     'GC and FC', '31 Oct 26'],
    ['Overhead keys', 'Events carries 87.5 per cent of premises on 13 per cent of revenue and '
                      'cannot be profitable on that key. Test it for FY28, do not change it '
                      'mid year.', 'GC and FC', '30 Nov 26'],
    ['Customers', 'Top three are 49 per cent of revenue. Agree a retention plan for Deloitte '
                  'and contract cover for Aware Super, Data#3 and Generation-e.',
     'GC and Sales', '14 Nov 26'],
    ['Cash', 'DSO 47.5 days. Getting to 30 releases 309k against net debt of 723k. Chase '
             'Property NSW 19.7k and Meet Magic 5.1k, both over 90 days.', 'FC', '24 Oct 26'],
])

out = f'{D}/CTS Monthly Business Review - September 2026.pptx'
prs.save(out)
print('saved', out)
