import json, pickle, re, sys
from collections import defaultdict
S = sys.argv[1]
tx, fin = pickle.load(open(S + '/data.pkl', 'rb'))
inv = json.load(open(S + '/inv_sep.json'))
byinv = defaultdict(list)
for f in fin: byinv[f['inv']].append(f)
EXTRA = {'INV-10652': 'ONSITE', 'INV-10669': 'VIDEO'}   # not on her Finance tab
alloc = defaultdict(float); notes = []
invtot = defaultdict(float)
for i in inv:
    n = i['invoice_number']
    lines = [(l['account_code'], float(l['line_amount_net'])) for l in i['line_items'] if l.get('account_code')]
    lines_d = [(l['account_code'], float(l['line_amount_net']), str(l['description'])) for l in i['line_items'] if l.get('account_code')]
    for a, v in lines: invtot[a] += v
    rows = byinv.get(n, [])
    ccs = {r['cc'] for r in rows}
    if not rows:
        cc = EXTRA[n]; notes.append((n, cc, 'not on Finance'))
        for a, v in lines: alloc[(a, cc)] += v
    elif len(ccs) == 1:
        cc = ccs.pop()
        for a, v in lines: alloc[(a, cc)] += v
    else:
        vid = sum(r['amt'] for r in rows if r['cc'] == 'VIDEO')
        got = 0
        for a, v, d in lines_d:
            isv = a == '42150' or (a == '42500' and d.strip() == 'Video')
            cc = 'VIDEO' if isv else 'PRODUCTION'
            alloc[(a, cc)] += v
            if isv: got += v
        assert abs(got - vid) < 0.01, (n, got, vid)
# journals and credit notes
def jcc(t):
    d = t['desc']
    if t['acct'] == '41175': return 'INTEGRATION'
    if t['src'] == 'Receivable Credit Note': return 'ONSITE'
    if 'Video Labour' in d or '[VID]' in d: return 'VIDEO'
    if 'DAFF' in d: return 'CONSULTING'
    if '[INT]' in d or 'Aware' in d: return 'INTEGRATION'
    return 'PRODUCTION'
for t in tx:
    if t['src'] != 'Receivable Invoice':
        alloc[(t['acct'], jcc(t))] += t['amt']
# check invoices vs account tabs
tabinv = defaultdict(float); tabtot = defaultdict(float)
for t in tx:
    tabtot[t['acct']] += t['amt']
    if t['src'] == 'Receivable Invoice': tabinv[t['acct']] += t['amt']
for a in sorted(set(tabinv) | set(invtot)):
    if abs(tabinv[a] - invtot[a]) > 0.005: print('INVOICE MISMATCH', a, round(tabinv[a], 2), round(invtot[a], 2))
res = {}
for a in sorted(tabtot):
    s = sum(v for (aa, c), v in alloc.items() if aa == a)
    res[a] = {c: round(v, 2) for (aa, c), v in alloc.items() if aa == a and abs(v) > 0.004}
    print(a, round(tabtot[a], 2), round(s, 2), res[a])
print(notes)
pickle.dump((res, {a: round(v, 2) for a, v in tabtot.items()}), open(S + '/res.pkl', 'wb'))
