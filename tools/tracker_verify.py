import openpyxl, datetime, re, zipfile
old = openpyxl.load_workbook('wb.xlsm', data_only=True)['Tracker']
nw  = openpyxl.load_workbook('FY27_Revenue_Tracker_v4.xlsx')
ds, dsh, ls = nw['Data'], nw['Dashboard'], nw['Lists']
def num(v): return v if isinstance(v,(int,float)) else 0
ok=True
def chk(label,a,b,tol=0.005):
    global ok; g=abs(a-b)<=tol; ok&=g
    print(f"  {'PASS' if g else 'FAIL'}  {label:<48} new={a:>14,.2f}  original={b:>14,.2f}")
def chkv(label,a,b):
    global ok; g=(a==b); ok&=g
    print(f"  {'PASS' if g else 'FAIL'}  {label:<48} {a}  (expected {b})")

DMAP={ls.cell(4+i,6).value: ls.cell(4+i,7).value for i in range(7)}
def repline(typ,dept): return 'WIP' if typ=='WIP' else DMAP.get(dept,'Unmapped')
def netrev(rl,amt,disc): return None if amt is None else (amt-num(disc) if rl=='Production' else amt)

print("=== A. Schema: one department input, one amount input ===")
hdrs=[ds.cell(1,c).value for c in range(1,41)]
chkv("no 'Source' column", 'Source' in hdrs, False)
chkv("'Department' appears once", hdrs.count('Department'), 1)
chkv("single amount input 'Amount'", hdrs.count('Amount'), 1)
chkv("'Invoice Value' gone", 'Invoice Value' in hdrs, False)
chkv("'WIP Movement' gone", 'WIP Movement' in hdrs, False)
chkv("'Reports as' present (calculated)", 'Reports as' in hdrs, True)

print("\n=== B. New rows auto-fill (calculatedColumnFormula) ===")
t=zipfile.ZipFile('FY27_Revenue_Tracker_v4.xlsx').read('xl/tables/table1.xml').decode()
ccf=re.findall(r'<tableColumn[^>]*name="([^"]+)"[^>]*>\s*<calculatedColumnFormula>',t)
print("  columns that auto-fill:", ccf)
for need in ['No.','Reports as','Net Revenue','Margin','Margin %','Date Check','Duplicate Check','Discounts as % of Net']:
    chkv(f"  auto-fills: {need}", need in ccf, True)

print("\n=== C. Totals tie back to the original workbook ===")
rows=[]
for r in range(2,ds.max_row+1):
    typ=ds.cell(r,2).value
    if not typ: continue
    rows.append(dict(typ=typ,date=ds.cell(r,3).value,dept=ds.cell(r,4).value,
                     job=ds.cell(r,8).value,amt=ds.cell(r,11).value,disc=ds.cell(r,30).value,
                     ch=ds.cell(r,32).value,li=ds.cell(r,33).value,
                     le=ds.cell(r,38).value,ei=ds.cell(r,39).value,se=ds.cell(r,40).value))
for x in rows: x['rl']=repline(x['typ'],x['dept']); x['net']=netrev(x['rl'],x['amt'],x['disc'])
chkv("record count",len(rows),181)
chkv("no Unmapped report lines",sum(1 for x in rows if x['rl']=='Unmapped'),0)
for rl,cnt in [('Support',35),('Production',49),('Consulting',32),('WIP',65)]:
    chkv(f"{rl} record count",sum(1 for x in rows if x['rl']==rl),cnt)
for rl,ocol in [('Support',74),('Production',119),('Consulting',134),('WIP',158)]:
    chk(f"{rl} Net Revenue",sum(num(x['net']) for x in rows if x['rl']==rl),
        sum(num(old.cell(r,ocol).value) for r in range(12,212)))
chk("Production Margin",sum(num(x['amt'])-(num(x['ch'])+num(x['li'])) for x in rows if x['rl']=='Production' and x['amt'] is not None),
    sum(num(old.cell(r,121).value) for r in range(12,212)))
chk("Consulting Margin",sum(num(x['amt'])-(num(x['le'])+num(x['ei'])+num(x['se'])) for x in rows if x['rl']=='Consulting' and x['amt'] is not None),
    sum(num(old.cell(r,141).value) for r in range(12,212)))

print("\n=== D. Dashboard months ===")
def bounds(i):
    m=7+i; y=2026+(m-1)//12; m=(m-1)%12+1
    m2,y2=(1,y+1) if m==12 else (m+1,y)
    return datetime.datetime(y,m,1), datetime.datetime(y2,m2,1)
specs=[('Support',74,69),('Production',119,97),('Consulting',134,125),('WIP',158,153)]
for i in range(12):
    a,b=bounds(i)
    for rl,ocol,odc in specs:
        exp=sum(num(old.cell(r,ocol).value) for r in range(12,212)
                if isinstance(old.cell(r,odc).value,datetime.datetime) and a<=old.cell(r,odc).value<b)
        got=sum(num(x['net']) for x in rows if x['rl']==rl
                and isinstance(x['date'],datetime.datetime) and a<=x['date']<b)
        if abs(exp)>0.005 or abs(got)>0.005: chk(f"{a.strftime('%b-%y')} {rl}",got,exp)

print("\n=== E. YOUR TEST: add a WIP record of 500 ===")
jul_before=sum(num(x['net']) for x in rows if x['rl']=='WIP'
               and isinstance(x['date'],datetime.datetime) and bounds(0)[0]<=x['date']<bounds(0)[1])
new=dict(typ='WIP',date=datetime.datetime(2026,7,15),dept='Production',job='TEST500',amt=500,disc=None)
new['rl']=repline(new['typ'],new['dept']); new['net']=netrev(new['rl'],new['amt'],new['disc'])
chkv("  Reports as",new['rl'],'WIP')
chkv("  Net Revenue",new['net'],500)
jul_after=jul_before+new['net']
print(f"  Jul-26 WIP before {jul_before:,.2f} -> after {jul_after:,.2f}  (moves by {jul_after-jul_before:,.2f})")
chkv("  Dashboard WIP moves by 500",round(jul_after-jul_before,2),500)
n2=dict(typ='WIP',date=datetime.datetime(2026,7,15),dept='Integration',job='T2',amt=-250,disc=None)
n2['rl']=repline(n2['typ'],n2['dept']); n2['net']=netrev(n2['rl'],n2['amt'],n2['disc'])
chkv("  negative WIP -250 also flows",n2['net'],-250)
n3=dict(typ='Invoice',date=datetime.datetime(2026,7,15),dept='Production',job='T3',amt=1000,disc=100)
n3['rl']=repline(n3['typ'],n3['dept']); n3['net']=netrev(n3['rl'],n3['amt'],n3['disc'])
chkv("  Production invoice 1000 less 100 discount",n3['net'],900)
n4=dict(typ='Invoice',date=datetime.datetime(2026,7,15),dept='Support',job='T4',amt=750,disc=None)
n4['rl']=repline(n4['typ'],n4['dept']); n4['net']=netrev(n4['rl'],n4['amt'],n4['disc'])
chkv("  Support invoice 750",n4['net'],750)

print("\n=== F. Formula function audit ===")
FNS=set()
for sh in nw.worksheets:
    for row in sh.iter_rows():
        for c in row:
            if isinstance(c.value,str) and c.value.startswith('='):
                FNS.update(re.findall(r'([A-Z][A-Z0-9_.]*)\s*\(',c.value))
BAD={'XLOOKUP','XMATCH','SORT','FILTER','UNIQUE','SEQUENCE','TEXTJOIN','CONCAT','IFS','SWITCH','MAXIFS','MINIFS'}
print("  used:",", ".join(sorted(FNS)))
chkv("no post-2007 / spilling functions",sorted(FNS&BAD),[])
print("\nRESULT:","ALL CHECKS PASSED" if ok else "*** FAILURES ***")
