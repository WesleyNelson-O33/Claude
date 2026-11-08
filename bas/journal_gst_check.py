import pickle,sys,re,collections,datetime as dt
js=pickle.load(open(sys.argv[1],'rb'))
q=[j for j in js if j['lines'] and dt.datetime(2026,7,1)<=j['lines'][0]['date']<=dt.datetime(2026,9,30)]
BANK={'11110','11120','11130','11131','11140'}
def amt(l): return round(l['dr']-l['cr'],2)
print("=== MANUAL JOURNALS touching GST 21310 ===")
man=[j for j in q if j['title'].rstrip().endswith('Manual') or '- Manual' in j['title'] or 'Rounding' in j['title']]
print('manual count',len(man))
for j in man:
    g=[l for l in j['lines'] if l['code']=='21310']
    if g: print(j['title'][:90], j['lines'][0]['date'].date(), [amt(l) for l in g])
print("=== SPEND/RECEIVE MONEY (bank, no AP/AR, not transfer/payment) ===")
for j in q:
    t=j['title']; codes={l['code'] for l in j['lines']}
    if re.match(r'ID \d+ (Payment|Reversed|Reversal|Bank Transfer)',t) or 'Manual' in t: continue
    if '21200' in codes or '11200' in codes: continue
    if not codes & BANK: continue
    gst=sum(amt(l) for l in j['lines'] if l['code']=='21310')
    base=sum(amt(l) for l in j['lines'] if l['code'] not in BANK and l['code']!='21310')
    accts=sorted({l['code']+' '+l['acct'] for l in j['lines'] if l['code'] not in BANK and l['code']!='21310'})
    r=(gst/base) if base else 0
    flag='' if abs(gst)<0.005 or abs(r-0.1)<0.003 else ' <-- GST not 10%'
    print(j['lines'][0]['date'].date(), t[:60], 'base',round(base,2),'gst',round(gst,2), accts[:4], flag)
print("=== CREDIT NOTES (AR credited w/o bank, or AP debited w/o bank) ===")
for j in q:
    t=j['title']; codes={l['code'] for l in j['lines']}
    if re.match(r'ID \d+ (Payment|Reversed|Reversal)',t) or codes & BANK: continue
    ar=sum(amt(l) for l in j['lines'] if l['code']=='11200')
    ap=sum(amt(l) for l in j['lines'] if l['code']=='21200')
    if ar<0 or ap>0:
        gst=sum(amt(l) for l in j['lines'] if l['code']=='21310')
        base=sum(amt(l) for l in j['lines'] if l['code'] not in ('11200','21200','21310'))
        print(j['lines'][0]['date'].date(), t[:70], 'AR',round(ar,2),'AP',round(ap,2),'base',round(base,2),'gst',round(gst,2), sorted({l['code'] for l in j['lines']}))
