import openpyxl, datetime, copy, re, sys
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter as L
sys.path.insert(0,'.')
from data_4nov import inv, bills, templates
D=datetime.datetime
wb=openpyxl.load_workbook('out/nov4_bank2.xlsx')
cf=wb['Cash Flow']; a=wb['Assumptions']; p=wb['Payroll']; rec=wb['Receivables']; pb=wb['Payables']; cr=wb['Creditors Run']
if 'Complete AV' in wb.sheetnames: del wb['Complete AV']
NOFILL=PatternFill(fill_type=None)
NW=8
FC=[L(2+2*i) for i in range(NW)]   # B D F H J L N P
AC=[L(3+2*i) for i in range(NW)]   # C E G I K M O Q
TF,TA,VAR,NOTE='R','S','T','U'
def nbd(d): return d if d.weekday()<5 else d+datetime.timedelta(days=7-d.weekday())

# ================= ASSUMPTIONS =================
a['B5']=D(2026,11,5); a['C5']='Thu 5 Nov. Eight weeks of Thu to Wed, the last one running to Thu 31 Dec. Week table is at the bottom of this tab.'
a['B6']='=29216.74+250411.39'; a['C6']='Cheque + savings at Wed 4 Nov, worked forward from the 1 Nov CommBiz statement with the 2 to 4 Nov items and the $150,000 of transfers. Replace with the CommBiz balances at 4 Nov.'
a['B7']=8; a['C7']='Each week has a Forecast and an Actual column on the Cash Flow tab. Monthly estimates are spread at 12/52 a week.'
a['A14']='Pay date - run 1 (fortnight Sat 31 Oct to Fri 13 Nov)'; a['B14']=D(2026,11,17); a['C14']='Tue 17 Nov (week 2). Super direct debits the day after.'
a['A15']='Pay date - run 2 (fortnight Sat 14 Nov to Fri 27 Nov)'; a['B15']=D(2026,12,1); a['C15']='Tue 1 Dec (week 4).'
a['A16']='Pay date - run 3 (fortnight Sat 28 Nov to Fri 11 Dec)'; a['B16']=D(2026,12,15); a['C16']='Tue 15 Dec (week 6).'
a['A17']='Payroll tax NSW - October return (due 7 Nov)'; a['A18']='Payroll tax other states - October (due 7 Nov)'
a['B19']=D(2026,11,6); a['C19']='7 Nov is a Saturday, so Fri 6 Nov (week 1).'
a['C21']='Q1 BAS is $115,000, paid 25 Nov through the agent (week 3). Set to 0 if it is paid outside this forecast.'
a['A39']='New invoices raised and collected inside the same week (estimate, from week 2)'; a['B39']=10000
a['C39']='October ran about $10k a week of small jobs billed and paid inside the month (Elkiem, Garvan, RSNSW, CBA foldback, APA). Nothing in week 1 because those are already on the Receivables tab. Drop it for the Christmas weeks if work stops.'
a['C40']='Oct was $411.39. Lands at month end - 30 Nov (week 4) and 31 Dec (week 8).'
for r in range(44,53):
    for c in 'ABC': a[f'{c}{r}'].value=None
for i in range(NW):
    r=44+i; a[f'A{r}']=i+1
    a[f'B{r}']='=Assumptions!$B$5' if i==0 else f'=C{r-1}+1'
    a[f'C{r}']=f'=B{r}+6' if i<NW-1 else f'=B{r}+7'
    for c in 'ABC': a[f'{c}{r}']._style=copy.copy(a[f'{c}44']._style) if i else a[f'{c}{r}']._style
a['C51']='=B51+7'
rows={53:('TAX AND PAY DATES INSIDE THIS FORECAST',None,None),
      54:('October PAYG withholding - monthly IAS','=B11*(101245.94+111704.13)','PAYG % x the two Oct runs ($212,950 gross, Xero cash P&L).'),
      55:('October IAS payment date',D(2026,11,20),'21 Nov is a Saturday, so Fri 20 Nov (week 3).'),
      56:('November PAYG withholding - monthly IAS','=B11*(115121.14+B9)','PAYG % x the 3 Nov run ($115,121) plus the 17 Nov run at the assumed gross.'),
      57:('November IAS payment date',D(2026,12,21),'Due 21 Dec, a Monday (week 7).'),
      58:('Payroll tax - November return (NSW + other states)','=B17+B18','Same total as October. Nov wages are a bit higher, so this is on the low side.'),
      59:('November payroll tax payment date',D(2026,12,7),'Mon 7 Dec (week 5).'),
      60:('Pay date - run 4 (fortnight Sat 12 Dec to Fri 25 Dec)',D(2026,12,29),'Tue 29 Dec (week 8). Christmas - it may get processed early, before the office closes.'),
      61:('Pay date - run 5 (fortnight Sat 26 Dec to Fri 8 Jan)',D(2027,1,12),'Tue 12 Jan - after this forecast.')}
for r,(x,y,z) in rows.items():
    a[f'A{r}']=x; a[f'B{r}']=y; a[f'C{r}']=z
    for c in 'ABC': a[f'{c}{r}']._style=copy.copy(a[f'{c}{42 if r==53 else 40}']._style)
    if isinstance(y,D): a[f'B{r}'].number_format=a['B25'].number_format
    elif isinstance(y,str): a[f'B{r}'].number_format=a['B22'].number_format

# ================= PAYROLL: four runs =================
p.insert_rows(8)
for c in range(1,11): p.cell(8,c)._style=copy.copy(p.cell(7,c)._style)
p['A8']='Run 4 - Sat 12 Dec to Fri 25 Dec'; p['B8']=D(2026,12,25); p['C8']='=Assumptions!$B$60'; p['D8']='=Assumptions!$B$9'
p['E8']='=D8*Assumptions!$B$10'; p['F8']='=D8*Assumptions!$B$11'; p['G8']='=D8*Assumptions!$B$12'; p['H8']='=Assumptions!$B$13'; p['I8']='=E8+G8+H8'
p['J8']='Paid Tue 29 Dec. Christmas week - likely processed early.'
p['A5']='Run 1 - Sat 31 Oct to Fri 13 Nov'; p['B5']=D(2026,11,13); p['J5']='Paid Tue 17 Nov.'
p['A6']='Run 2 - Sat 14 Nov to Fri 27 Nov'; p['B6']=D(2026,11,27); p['J6']='Paid Tue 1 Dec. Last run of November - accrue bonus leave (Checklist step 48).'
p['A7']='Run 3 - Sat 28 Nov to Fri 11 Dec'; p['B7']=D(2026,12,11); p['J7']='Paid Tue 15 Dec.'
for c in 'DEFGHI': p[f'{c}9']=f'=SUM({c}5:{c}8)'
p['A10']='Payroll tax inside this forecast'; p['A11']='October return - NSW + other states, paid Fri 6 Nov'; p['B11']='=Assumptions!$B$17+Assumptions!$B$18'; p['C11']='=Assumptions!$B$19'
p['A12']='November return - paid Mon 7 Dec'; p['B12']='=Assumptions!$B$58'; p['C12']='=Assumptions!$B$59'
p['A13']='Total payroll tax'; p['B13']='=B11+B12'
p['A14']='Payroll tax on Dec wages is due 7 Jan - after this forecast.'
for r in range(18,26):
    if isinstance(p[f'E{r}'].value,str) and p[f'E{r}'].value.startswith('=IF('): p[f'E{r}']=f'=IF(C{r}=0,0,D{r}/C{r})'
for r in range(18,32):
    v=p[f'A{r}'].value
    if v=='Average gross per fortnight (last two runs)': p[f'B{r}']=f'=(C{r-1}+115121.14)/2'
    if isinstance(p[f'D{r}'].value,str) and p[f'D{r}'].value.startswith('=B'): p[f'D{r}']=f'=B{r}+C{r}'

# ================= CASH FLOW: widen to eight weeks =================
for m in list(cf.merged_cells.ranges): cf.unmerge_cells(str(m))
cf.insert_cols(12,6)
for r in range(1,73):
    for i in range(5,NW):
        cf[f'{FC[i]}{r}']._style=copy.copy(cf[f'J{r}']._style); cf[f'{AC[i]}{r}']._style=copy.copy(cf[f'K{r}']._style)
cf.column_dimensions[NOTE].width=64; cf.column_dimensions['O'].width=13
for col in FC+AC+[TF,TA,VAR]: cf.column_dimensions[col].width=13
cf.column_dimensions['A'].width=46
# headers
for i in range(NW):
    cf[f'{FC[i]}5']=f'Week {i+1}'; cf[f'{AC[i]}5']=None
    cf[f'{FC[i]}6']=f'=TEXT(Assumptions!$B${44+i},"d mmm")&" - "&TEXT(Assumptions!$C${44+i},"d mmm")'; cf[f'{AC[i]}6']=None
    cf[f'{FC[i]}7']='Forecast'; cf[f'{AC[i]}7']='Actual'
    cf.merge_cells(f'{FC[i]}5:{AC[i]}5'); cf.merge_cells(f'{FC[i]}6:{AC[i]}6')
    cf[f'{AC[i]}8']=0
cf[f'{TF}5']='Total'; cf[f'{TA}5']=None; cf.merge_cells(f'{TF}5:{TA}5'); cf[f'{VAR}5']='Variance'; cf[f'{NOTE}5']='Notes'
cf[f'{TF}6']=None; cf[f'{VAR}6']='(weeks with'; cf[f'{TF}7']='Forecast'; cf[f'{TA}7']='Actual'; cf[f'{VAR}7']='actuals)'
cf.merge_cells(f'A3:{NOTE}3')
for r in (69,70,71,72): cf.merge_cells(f'A{r}:{NOTE}{r}')
# opening
for i in range(NW):
    f,ac=FC[i],AC[i]
    cf[f'{f}10']='=Assumptions!$B$6' if i==0 else f'=IF({AC[i-1]}$8=1,{AC[i-1]}59,{FC[i-1]}58)'
    cf[f'{ac}10']=f'={f}10'
cf[f'{TF}10']='=B10'; cf[f'{TA}10']='=C10'
# line formulas: take week 1 formula text and shift the week table row
line_rows=[13,14,19,20,21,22,23,27,28,29,30,31,35,36,37,38,39,40,41,42,43,44,45,46,47,48,52]
base={r:cf[f'B{r}'].value for r in line_rows}
base[13]=base[13]  # SUMIFS on receivables
base[14]='=0'
base[21]='=IF(AND(Assumptions!$B$55>=Assumptions!$B$44,Assumptions!$B$55<=Assumptions!$C$44),Assumptions!$B$54,0)+IF(AND(Assumptions!$B$57>=Assumptions!$B$44,Assumptions!$B$57<=Assumptions!$C$44),Assumptions!$B$56,0)'
base[22]='=IF(AND(Assumptions!$B$19>=Assumptions!$B$44,Assumptions!$B$19<=Assumptions!$C$44),Assumptions!$B$17+Assumptions!$B$18,0)+IF(AND(Assumptions!$B$59>=Assumptions!$B$44,Assumptions!$B$59<=Assumptions!$C$44),Assumptions!$B$58,0)'
for r in line_rows:
    t=base[r]
    if not isinstance(t,str): continue
    t=t.replace('Payroll!$C$5:$C$7','Payroll!$C$5:$C$8').replace('Payroll!$E$5:$E$7','Payroll!$E$5:$E$8').replace('Payroll!$G$5:$G$7','Payroll!$G$5:$G$8').replace('Payroll!$H$5:$H$7','Payroll!$H$5:$H$8')
    t=re.sub(r'/5\b','*12/52',t)
    base[r]=t
for i in range(NW):
    f=FC[i]; w=44+i
    for r in line_rows:
        t=base[r]
        if not isinstance(t,str): continue
        t=t.replace('$B$44',f'$B${w}').replace('$C$44',f'$C${w}')
        if r==13 and i>0: t=t+'+Assumptions!$B$39'
        if r==14: t='=Assumptions!$B$40' if i in (3,7) else '=0'
        cf[f'{f}{r}']=t
        cf[f'{AC[i]}{r}']=None
    cf[f'{TF}{r}']=None
for r in line_rows:
    cf[f'{TF}{r}']='=SUM('+','.join(f'{c}{r}' for c in FC)+')'
    cf[f'{TA}{r}']='=SUM('+','.join(f'{c}{r}' for c in AC)+')'
    cf[f'{VAR}{r}']='='+'+'.join(f'({AC[i]}{r}-{FC[i]}{r})*{AC[i]}$8' for i in range(NW))
# subtotals and totals
sub={15:'={c}13+{c}14',24:'={c}19+{c}20+{c}21+{c}22+{c}23',32:'={c}27+{c}28+{c}29+{c}30+{c}31',
     49:'={c}35+{c}36+{c}37+{c}38+{c}39+{c}40+{c}41+{c}42+{c}43+{c}44+{c}45+{c}46+{c}47+{c}48',53:'={c}52',55:'={c}24+{c}32+{c}49+{c}53',57:'={c}15-{c}55'}
for r,t in sub.items():
    for c in FC+AC+[TF,TA,VAR]: cf[f'{c}{r}']=t.format(c=c)
for i in range(NW):
    cf[f'{FC[i]}58']=f'={FC[i]}10+{FC[i]}57'; cf[f'{AC[i]}58']=f'=IF({AC[i]}$8=1,{AC[i]}10+{AC[i]}57,"")'
    cf[f'{AC[i]}59']=None; cf[f'{AC[i]}60']=f'=IF({AC[i]}$8=1,{AC[i]}58-{AC[i]}59,"")'
    cf[f'{FC[i]}63']=f'={FC[i]}19+{FC[i]}20+{FC[i]}23'; cf[f'{FC[i]}64']=f'=IF(Payroll!$I$5=0,0,{FC[i]}58/(Payroll!$I$5/2))'
    cf[f'{AC[i]}63']=None; cf[f'{AC[i]}64']=None
cf[f'{TF}58']=f'={FC[-1]}58'
nested='""'
for ac in AC: nested=f'IF({ac}$8=1,{ac}58,{nested})'
cf[f'{TA}58']='='+nested
cf[f'{VAR}58']=None
for c in [TF,TA,VAR]:
    for r in (59,60,63,64): cf[f'{c}{r}']=None
cf[f'{TF}65']=f'=SUMIFS(Receivables!$F$5:$F$65,Receivables!$H$5:$H$65,">"&Assumptions!$C${43+NW})'
cf[f'{TF}66']=f'=SUMIFS(Payables!$D$19:$D$95,Payables!$K$19:$K$95,"Yes",Payables!$E$19:$E$95,"<>Card",Payables!$G$19:$G$95,">"&Assumptions!$C${43+NW})'
cf['A65']='   Receipts expected after 31 Dec (not in this forecast)'; cf['A66']='   Payments held past 31 Dec (not in this forecast)'
cf['A2']='Weekly Cash Flow Forecast vs Actual - 5 November to 31 December 2026'
cf['A3']=("How this works: the Forecast columns are formulas off the Receivables, Payables, Creditors Run, Payroll and Assumptions tabs. Each week, put the actual movements in the yellow Actual column, type in the closing bank balance from CommBiz (cheque + savings), and set 'Actuals entered' to 1. "
          "The next week then starts from the real bank balance and the difference row shows anything that's been missed. Monthly estimates are spread at 12/52 a week.")
cf.freeze_panes='B9'
notes={
 8:'Set to 1 once the week is done and the CommBiz balance is in. Then the next week opens from the real number.',
 13:'Receivables tab plus $10k a week of small jobs billed and paid inside the week from week 2 (Assumptions B39). Week 1: Deloitte Oct $117k, Bankwest, CBA Brisbane and the overdue small ones. Week 2: CBA production kit $118k plus the CBA Oct contract invoices. Week 4: ASX AGM $70.6k, TLC $39k, April 5 $32k, PwC, Aware, Bankwest Nov. Week 5: Cochlear $70k, Deloitte Nov. Week 6: CBA Nov contract invoices. Week 8: PwC and Aware Nov.',
 14:'Savings interest at month end - 30 Nov (week 4) and 31 Dec (week 8).',
 19:'Payroll tab - gross x net %. Fortnight is Sat to Fri and the money goes out the Tuesday after: 17 Nov (week 2), 1 Dec (week 4), 15 Dec (week 6) and 29 Dec (week 8). The Christmas one may go early.',
 20:'Gross x 12%. Direct debits the day after pay day.',
 21:'Monthly IAS: Oct PAYG ~$45k on Fri 20 Nov (week 3) and Nov PAYG ~$48k on Mon 21 Dec (week 7). Assumptions rows 54 to 57.',
 22:'Oct wages $8,603 on Fri 6 Nov (week 1) and the Nov return on Mon 7 Dec (week 5), same total used.',
 23:'Placeholder - goes out with the ABA.',
 27:'Unbilled estimate only. The Contact Group fitout and anything else already in Xero is in the creditors run.',
 28:'Estimate, $40k a month spread at 12/52 a week. Sep was ~$39k, Oct ~$27k.',
 29:'Estimate, spread evenly. The hire bills already in Xero (WordBirds, Riedel, 2020 Events, Headroom, Start Webcast) are in the creditors run.',
 30:'LiveU and Spacera templates for Nov and Dec plus the unbilled estimate. Appspace is in the creditors run.',
 31:'Estimate plus the Goget template (card).',
 35:'Dec rent ~$28,896 billed 1 Dec, paid on the day like Oct and Nov (week 4). Jan rent falls after 31 Dec.',
 36:'First Focus CORE + Azure, DD 16 Nov (week 2) and 16 Dec (week 6).',
 37:'RJW digital marketing $4,675 on 23 Nov and 21 Dec.',
 38:'8020 Advisors $3,850 and Employsure $880 each month, plus the estimate. Easton Belle legal was paid 2 Nov.',
 39:'CEO Institute $2,310 (7 Dec) and AVIXA $2,850 (11 Dec). Card subscriptions sit in the card line.',
 40:'Optus mobiles Oct (9 Nov) and Nov (9 Dec), Optus home internet, Telstra Nighthawk - BPAY.',
 41:'Estimate. Seek ads go on the card.',
 42:'Brookvale was paid 2 Nov. Nothing else expected - estimate is nil.',
 43:'Ricoh DD and Posh each month, plus the estimate.',
 44:'Energy Australia DDs the PL card so it shows in the credit card line, not here.',
 45:'Estimate - bank charges and Stripe.',
 46:'~$25k a month transferred to the PL and DL cards, spread at 12/52 a week. Covers every Card row on the Payables tab.',
 47:'Nothing listed.',
 48:'Creditors Run tab - the 29 open bills at 4 Nov ($84.6k), each on its due date. Appspace $19k and Start Webcast $12.8k go Fri 6 Nov. The Dec run is Tue 1 Dec but nothing is sitting in it yet - December bills will land as they come in.',
 52:'Q1 BAS $115,000 on Wed 25 Nov through the agent (week 3). The Q2 BAS is due 28 Feb, after this forecast.',
 58:"Opening + net. The Actual column fills in once the week's flag is 1, and the total column shows the latest completed week.",
 59:'Type the real closing balance (cheque + savings) in at the end of each week. The opening balance on the Assumptions tab is worked forward from the 1 Nov statement - replace it with the CommBiz balances at 4 Nov first.',
 60:"Should be nil once everything's in. This is a sense check.",
 65:'Dec contract invoices (raised 31 Dec, paid Jan), Meet Magic, Jaaz.',
 66:'Jan rent, Spacera Dec, Optus mobiles Dec, Posh Dec.',
}
for r,v in notes.items(): cf[f'{NOTE}{r}']=v
cf[f'{NOTE}6']=None; cf[f'{NOTE}7']=None
cf['A69']="1. Tax is the big block - Oct payroll tax $8.6k (6 Nov), Oct IAS ~$45k (20 Nov), Q1 BAS $115k (25 Nov), Nov payroll tax $8.6k (7 Dec) and Nov IAS ~$48k (21 Dec). About $225k out between now and Christmas, with the BAS week (week 3) the heaviest."
cf['A70']="2. CBA Place South $118,250 (production kit, due 15 Nov) is the biggest receipt and is not in yet - week 2. Deloitte's Oct ($117k) and Nov (~$114k) contract invoices are the other two big ones. If CBA pay on the due date the week 2 dip is fine; if they go past it, week 3 with the BAS gets tight."
cf['A71']="3. Overdue debtors: Downer $39,664 (six weeks), Microsoft $13,846 (two months), Ventia $10,016, ASX $5,748, ON24 $3,006 and a few small new customers - about $85k. I have them landing in weeks 2 to 3. If none of it comes, closing is ~$85k lower. Chase Downer and Microsoft this week."
cf['A72']="4. December is thin on receipts - the Nov contract invoices land mid-December and the AGM season is over, so December trading is the $10k a week estimate. Four pay runs sit in the eight weeks (~$105k each with super), the last one in Christmas week. Cochlear's $70k (due 4 Dec) and the ASX AGM $70.6k (30 Nov) are the ones to watch."

# ================= RECEIVABLES =================
for r in range(5,66):
    for c in range(1,12):
        cell=rec.cell(r,c); cell.value=None; cell.fill=NOFILL
base_st=[copy.copy(rec.cell(36,c)._style) for c in range(1,12)]
def rrow(r,vals,note):
    for c in range(1,12): rec.cell(r,c)._style=copy.copy(base_st[c-1])
    for c,v in zip('ABCDEFGH',vals): rec[f'{c}{r}']=v
    rec[f'I{r}']=f'=H{r}-E{r}'
    rec[f'J{r}']=f'=IF(H{r}<Assumptions!$B$5,"Before",IF(H{r}>Assumptions!$C${43+NW},"After",INT((H{r}-Assumptions!$B$5)/7)+1))'
    rec[f'K{r}']=note
inv2=[]
for row in inv:
    row=list(row)
    if row[1]=='INV-10464': row[7]=D(2027,1,29); row[8]='DOUBTFUL - four months overdue. Left out of the forecast. Needs escalating.'
    if row[1]=='INV-10621': row[7]=D(2027,1,29); row[8]='Staff recovery - goes through payroll, no cash. Left out.'
    if row[1]=='INV-10725': row[8]='Cochlear paid the Aug results webcast a day before due - Thu 3 Dec (week 5).'
    inv2.append(tuple(row))
tbr=[
 ('Deloitte Services Trust','TBR','Nov contract invoice (template $113,988)',D(2026,11,30),D(2026,12,30),113987.99,'To be raised',D(2026,12,9),'Raised 30 Nov. Deloitte pay about a week after - Wed 9 Dec (week 5).'),
 ('Deloitte Services Trust','TBR','Nov additionals (estimate)',D(2026,11,30),D(2026,12,30),5000.0,'To be raised',D(2026,12,9),'Oct additionals were only $1,187, Sep $18,892. Call it $5k.'),
 ('Commonwealth Bank (CBA Place South)','TBR','Nov contract + dedicated tech (Oct actual $28.7k)',D(2026,11,30),D(2026,12,30),28699.11,'To be raised',D(2026,12,15),'CBA pay about 15 days after invoice - Tue 15 Dec (week 6).'),
 ('Commonwealth Bank (CBA Place South)','TBR','Nov live event support (estimate)',D(2026,11,30),D(2026,12,30),8000.0,'To be raised',D(2026,12,15),'Oct was $11,025, Sep $5,398.'),
 ('Commonwealth Bank (CBA Square)','TBR','Nov contract (Oct actual $14.4k)',D(2026,11,30),D(2026,12,30),14384.7,'To be raised',D(2026,12,15),'Same as Place South.'),
 ('Commonwealth Bank (CBA Melbourne)','TBR','Nov contract (Oct actual $15.8k)',D(2026,11,30),D(2026,12,30),15774.0,'To be raised',D(2026,12,15),'Same.'),
 ('Commonwealth Bank (CBA Brisbane)','TBR','Nov contract (Oct actual $8.2k)',D(2026,11,30),D(2026,12,30),8177.95,'To be raised',D(2026,12,7),'A week after invoice - Mon 7 Dec (week 5).'),
 ('Bankwest','TBR','Nov contract + additionals (Oct actual $14.7k)',D(2026,11,30),D(2026,12,30),14740.0,'To be raised',D(2026,12,2),'Bankwest pay within days - Wed 2 Dec (week 4).'),
 ('PwC Services','TBR','Nov contract x 2 (PO2601702104 + PO2601702371)',D(2026,11,30),D(2026,12,30),20962.7,'To be raised',D(2026,12,30),'PwC pay on the due date - Wed 30 Dec (week 8).'),
 ('Aware Super','TBR','Annual National Maintenance - November',D(2026,11,30),D(2026,12,30),12100.0,'To be raised',D(2026,12,28),'Aware pay a few days before due - Mon 28 Dec (week 8).'),
]
r=5
for row in inv2+tbr: rrow(r,row[:8],row[8]); r+=1
rec['A2']=("Xero aged receivables at 4 Nov ($726,549.38, 56 invoices - the small CBA ones are grouped) plus the November contract invoices still to be raised (October actuals used as the estimate). "
           "Expected date (blue) is my call based on how each customer has been paying - change it and the forecast moves. December contract invoices are raised 31 Dec and paid in January, so they are not here.")
rec['A67']='of which open in Xero at 4 Nov (should be $726,549.38)'
rec['A68']='of which expected inside the forecast (5 Nov to 31 Dec)'; rec['F68']=f'=SUMIFS(F5:F65,H5:H65,">="&Assumptions!$B$5,H5:H65,"<="&Assumptions!$C${43+NW})'
rec['A69']='of which pushed past 31 Dec (Meet Magic, Jaaz)'

# ================= PAYABLES =================
tmpl=[copy.copy(pb.cell(60,c)._style) for c in range(1,16)]
for r in range(19,96):
    for c in range(1,16):
        cell=pb.cell(r,c); cell.value=None; cell._style=copy.copy(tmpl[c-1])
    pb[f'H{r}']=f'=IF(G{r}<Assumptions!$B$5,"Before",IF(G{r}>Assumptions!$C${43+NW},"After",INT((G{r}-Assumptions!$B$5)/7)+1))'
    pb[f'K{r}']=f'=IF(AND(\'Creditors Run\'!$B$5=1,I{r}="Xero bill",LEFT(J{r},4)<>"PAID"),"No - in the creditors run","Yes")'
dec_templates=[
 ('Employsure','44 of 60 - December',D(2026,12,5),880.0,'EFT','Professional fees','Repeating template','Template, due Sat 5 Dec so Mon 7 Dec.'),
 ('Cleveland Partners (8020 Advisors)','Monthly advisory fees - December',D(2026,12,9),3850.0,'EFT','Professional fees','Repeating template','Template $3,850 a month.'),
 ('Optus','INV - staff mobiles Nov',D(2026,12,9),804.06,'BPAY','Telephone & internet','Repeating template','BPAY on the due date.'),
 ('Posh Services (Ambius)','Indoor plants - November',D(2026,12,9),245.0,'EFT','Office','Repeating template','On the due date.'),
 ('The CEO Institute','Executive membership (JS) - quarterly',D(2026,12,7),2310.0,'EFT','Dues & subscriptions','Repeating template','Template, due 7 Dec.'),
 ('AVIXA','DL annual silver membership',D(2026,12,11),2850.0,'EFT','Dues & subscriptions','Repeating template','Template, due 11 Dec.'),
 ('Telstra','Recharges x 4 - December',D(2026,12,9),176.0,'Card','Telephone & internet','Repeating template','PL card.'),
 ('LiveU','SOLO standard data pack - December',D(2026,12,14),1485.0,'EFT','Client subscriptions','Repeating template','On the due date.'),
 ('First Focus IT','CORE Intelligent Managed Services - December',D(2026,12,16),13755.77,'DD','IT support','Repeating template','DD on the due date.'),
 ('First Focus IT','Managed Services Azure - December',D(2026,12,16),710.5,'DD','IT support','Repeating template','DD on the due date.'),
 ('Insphire (Current RMS)','December',D(2026,12,16),431.0,'Card','Dues & subscriptions','Repeating template','PL card.'),
 ('Mailchimp','December',D(2026,12,18),37.62,'Card','Dues & subscriptions','Repeating template','PL card.'),
 ('Optus','Lugstein home internet - December',D(2026,12,20),101.2,'BPAY','Telephone & internet','Repeating template','Due Sun 20 Dec so Mon 21 Dec.'),
 ('RJW Group Holdings','Digital growth marketing - December',D(2026,12,21),4675.0,'EFT','Advertising','Repeating template','On the due date.'),
 ('Telstra','Nighthawk data plans - December',D(2026,12,21),130.0,'BPAY','Telephone & internet','Repeating template','On the due date.'),
 ('Optus','Office internet & landline - December',D(2026,12,21),629.66,'Card','Telephone & internet','Repeating template','PL card.'),
 ('Xero / Adobe / Spotify / Avangate / misc','Small monthly subscriptions - December',D(2026,12,24),700.0,'Card','Dues & subscriptions','Repeating template','PL / DL cards.'),
 ('Zoom','Monthly subscription - December',D(2026,12,28),323.97,'Card','Dues & subscriptions','Repeating template','PL card.'),
 ('Employment Hero','Subscription - November',D(2026,12,29),2048.2,'Card','Dues & subscriptions','Repeating template','PL card - in the card line.'),
 ('Goget Carshare','November usage',D(2026,12,29),1783.07,'Card','Direct travel & freight','Repeating template','PL card - in the card line.'),
 ('Ricoh Australia','Printing Nov - DD',D(2026,12,31),235.6,'DD','Office','Repeating template','Direct debit, end of month.'),
 ('Investa Asset Management','Rent January 2027',D(2027,1,1),28895.72,'EFT','Rent','Repeating template','Billed 1 Jan - after this forecast.'),
 ('SPACERA','Aware Super national support - December',D(2027,1,1),5728.8,'EFT','Client subscriptions','Repeating template','Due 1 Jan - after this forecast.'),
 ('Optus','Staff mobiles - December',D(2027,1,9),804.06,'BPAY','Telephone & internet','Repeating template','After this forecast.'),
]
nov_templates=[t for t in templates if t[2]<D(2026,12,2) or t[0] in ('SPACERA','Investa Asset Management')]
nov_templates=[t for t in nov_templates if not (t[0]=='Posh Services (Ambius)' or t[0]=='The CEO Institute' or t[0]=='AVIXA' or (t[0]=='Optus' and 'Staff mobiles' in t[1]))]
r=19
for b in bills:
    sup,ref,due,amt,meth,cat,note,net,gst,invd,payon=b
    for c,v in zip('ABCDEFGIJ',(sup,ref,due,amt,meth,cat,payon,'Xero bill',note)): pb[f'{c}{r}']=v
    r+=1
for t in nov_templates+dec_templates:
    sup,ref,due,amt,meth,cat,src,note=t
    for c,v in zip('ABCDEFGIJ',(sup,ref,due,amt,meth,cat,nbd(due),src,note)): pb[f'{c}{r}']=v
    r+=1
print('payables rows to',r-1)
pb['A97']='TOTAL bills listed from Xero (aged payables at 4 Nov, should be $84,568.84)'
pb['A2']=("Xero aged payables at 4 Nov ($84,568.84, 29 bills) plus the repeating bill templates for November and December. Card rows get paid on the PL/DL cards and sit in the credit card line on the Cash Flow tab. "
          "With the Creditors Run tab switched on, the open Xero bills come through the Creditors run row and the templates come through the category lines.")
pb['A16']='of which agreed / due this week'
# Creditors Run already holds the 29 bills at 4 Nov with the 1 Dec run date - keep as is
cr['C4']='Tue 1 Dec - the next monthly run (week 4). Anything with its own date in column L ignores this, and everything open at 4 Nov is dated to its due date. Paste the December aged payables in when you run it.'
wb.active=0
wb.calculation.fullCalcOnLoad=True
wb.save('out/yearend.xlsx'); print('saved')
