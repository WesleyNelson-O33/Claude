import pickle,sys,re,collections,datetime as dt
js=pickle.load(open(sys.argv[1],'rb'))
rev=set()
for j in js:
    m=re.search(r'Reversal of ID (\d+)',j['title'])
    if m: rev.add(int(m.group(1)))
byid={j['lines'][0]['id']:j for j in js if j['lines']}
def amt(l): return round(l['dr']-l['cr'],2)
for i in [69149,69407,70627,70629,69840,70433,69661,69631,70745,70741]:
    j=byid.get(i); 
    if not j: print(i,'missing');continue
    print(i,'REVERSED' if i in rev else 'LIVE',j['title'][:70],'posted',j['lines'][0]['posted'])
    for l in j['lines']: print('   ',l['date'].date(),l['code'],l['acct'],amt(l))
# 21420 and 21310 movement by month (live = all, reversals net)
for code in ('21310','21420'):
    m=collections.defaultdict(float)
    for j in js:
        for l in j['lines']:
            if l['code']==code: m[l['date'].strftime('%Y-%m')]+=amt(l)
    print(code,{k:round(v,2) for k,v in sorted(m.items())})
# ATO payments all year
for j in js:
    if 'Australian Taxation Office' in j['title'] and j['lines'][0]['id'] not in rev and not j['title'].startswith('ID %d Reversed'%j['lines'][0]['id']):
        print(j['lines'][0]['date'].date(), j['title'][:50], [(l['code'],amt(l)) for l in j['lines']])
