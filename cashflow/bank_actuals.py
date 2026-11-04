import openpyxl, datetime, copy, re, sys
from openpyxl.styles import PatternFill, Font
sys.path.insert(0,'/home/user/Claude/cashflow'); sys.path.insert(0,'.')
import inject_values
from data_4nov import inv, bills, templates
D=datetime.datetime
SRC='in2/her29.xlsx'
# 1. freeze the 28 Sep forecast columns as values (those weeks are over)
vals=inject_values._values(SRC)
wb=openpyxl.load_workbook(SRC)
cf=wb['Cash Flow']; rec=wb['Receivables']; pb=wb['Payables']; cr=wb['Creditors Run']; a=wb['Assumptions']
for col in 'BDFHJ':
    for r in list(range(13,15))+list(range(19,24))+list(range(27,32))+list(range(35,49))+[52]:
        c=cf[f'{col}{r}']
        if isinstance(c.value,str) and c.value.startswith('='):
            v=vals.get(('CASH FLOW',f'{col}{r}'))
            c.value=round(float(v),2) if isinstance(v,(int,float)) else 0
cf['A3']=("How this works: the Forecast columns are the 28 Sep forecast, frozen as numbers now that the five weeks are over. The Actual columns are straight off the CommBiz cheque account statement (2 Aug to 1 Nov) plus Xero for 2 to 4 Nov. "
          "Transfers between cheque and savings are left out because the bank line is cheque + savings. The Receivables, Payables and Creditors Run tabs show Xero at 4 Nov.")
NOFILL=PatternFill(fill_type=None)
def put(col,rows):
    for r,v in rows.items():
        cf[f'{col}{r}']=v; cf[f'{col}{r}']._style=copy.copy(cf['C13']._style)
blank={r:0 for r in [13,14,19,20,21,22,23,27,28,29,30,31,35,36,37,38,39,40,41,42,43,44,45,46,47,48,52]}
# ---- week 1: 28 Sep - 7 Oct ----
w={**blank,13:'=41925.4+6172.39+1092.3+1975.51+58024.76+13046.62+4352.45+13156+9350+93.5+627+8043.2',14:430.17,19:79736.05,20:14446.48,
   27:139086.23,38:880,43:235.6,45:10,46:10000,48:77302.54,59:'=21747.4+250000'}
put('C',w)
# ---- week 2: 8 - 14 Oct ----
w={**blank,13:'=374+131574.88+1661+506+9086.39+4908.2+19733.45+237987.4+748',19:'=179.86+2584.06',20:'=32.14+412.09',22:5918.58,
   29:1488.3,36:692.84,40:'=804.11+136',48:86510.67,59:'=29568.07+550000'}
put('E',w)
# ---- week 3: 15 - 21 Oct ----
w={**blank,13:'=13136.2+57591.56+2448.88+5398.25+2832.5+23317.11+8635.91',19:83859.74,20:15133.2,36:13566,40:101.2,45:'=0.83+36.03',46:10293.26,48:112761.41,59:'=57176.81+400000'}
put('G',w)
# ---- week 4: 22 - 28 Oct ----
w={**blank,13:'=16090.8+3169.1',23:1989.31,27:12540,38:342,47:'=1862.63+54.12',48:21929.13,59:'=37719.52+400000'}
put('I',w)
# ---- week 5: 29 Oct - 4 Nov (statement to 1 Nov, Xero and the transfers keyed for 2-4 Nov) ----
w={**blank,13:'=20962.7+1034+3498+1116.5+26660.69+11850.16+1969',14:411.39,19:90993.53,20:15994.3,35:28895.72,45:10,48:'=66238.98+10555.92+5728.8+3932.5+2585+257.4+210.2',
   59:'=29408.22+250411.39'}
put('K',w)
for col in 'CEGIK': cf[f'{col}8']=1
notes={
 'O8':'All five weeks are done (flag 1). Weeks 1 to 4 tie to the CommBiz cheque statement to the cent. Week 5 is the statement to 1 Nov plus Xero and the savings transfers for 2 to 4 Nov.',
 'O13':'Straight off the cheque statement, listed in the cell. Week 2 has Aware $237,987 (13 Oct), Deloitte $131,575 and Property NSW $19,733. Week 3 is the CBA Sep invoices $57,592 and Aware Appspace. Week 5 has Ventia $26,661, Gen-e $11,850 and PwC $20,963.',
 'O14':'Savings interest - $430.17 on 30 Sep and $411.39 on 30 Oct. It is credited to savings, so it is in the bank line. Transfers between cheque and savings are not income and are left out.',
 'O19':'CTSWAGES from the statement: 6 Oct $79,736.05, two top-ups 8 and 9 Oct ($2,764), 20 Oct $83,859.74, 3 Nov $90,993.53.',
 'O20':'HeroClear (Employment Hero) direct debits off the statement: 7 Oct $14,446.48, 9 and 13 Oct $444.23, 20 Oct $15,133.20, then $15,994.30 for the 3 Nov run.',
 'O21':'Nothing in Oct - Jul and Aug were paid 21 Aug and 21 Sep, Sep goes on the Q1 BAS.',
 'O22':'Revenue NSW $5,918.58 went out 8 Oct (week 2), not 7 Oct. The other states ($2,684) are not on the cheque statement - check if they went on the card or are still owing.',
 'O23':'DL CTS EXP SEPT $1,989.31 on 27 Oct is DL expenses. The per diems go out inside the wages.',
 'O27':'$139,086.23 Complete AV materials 29 Sep. Week 4 is SOR-1423275 $12,540 on 27 Oct - matches the equipment purchase in Xero, check what it was. The 8 Oct Complete AV labour ($39,033) is inside the 8 Oct batch on the creditors run line.',
 'O28':'Estimate - not paid separately. Any sub-contractors paid went out in the batches on the creditors run line.',
 'O29':'Videocraft $1,488.30 on 12 Oct. The rest of the hires went out in the batches.',
 'O30':'Nothing paid on its own - Eptura and LiveU were in the 8 Oct batch.',
 'O31':'Nothing paid on its own - any freight is inside the batches or on the card.',
 'O35':'Nov rent $28,895.72 went out 2 Nov (Xero). Sep and Oct rent were paid before this forecast.',
 'O36':'First Focus Azure $692.84 on 14 Oct and CORE $13,566 on 15 Oct, both direct debits.',
 'O37':'RJW went out inside the 23 Oct batch.',
 'O38':'Employsure (Peninsula) $880 on 6 Oct, ASIC $342 on 23 Oct.',
 'O39':'Nothing on the cheque account - card subscriptions sit in the card line.',
 'O40':'Optus mobiles $804.11 and Telstra $136 on 8 Oct, Optus home internet $101.20 on 16 Oct - all BPAY.',
 'O41':'Seek ads went on the card.',
 'O42':'Brookvale $66,239 went out in the 2 Nov creditors run (week 5), not week 2 as forecast.',
 'O43':'Ricoh DD $235.60 on 30 Sep. Posh went in the 8 Oct batch, Ricoh Sep in the 2 Nov run.',
 'O44':'Energy Australia DDs the PL card so it shows in the credit card line, not here.',
 'O45':'Account fee $10 on the 1st and CommBiz transfer fees 15 Oct.',
 'O46':'Transfers to the PL card: $10,000 on 7 Oct and $10,293.26 on 21 Oct.',
 'O47':'CTS ad-hoc payments 23 Oct $1,862.63 and $54.12 - check what they were.',
 'O48':'The payment batches off the statement: 1 Oct $77,302.54 (the creditors run), 8 Oct $86,510.67 (Complete AV labour $39,033, LiveU $26,004, Amora, New Office, Eptura, Valiant and the small ones), 16 Oct $112,761.41, 23 Oct $21,929.13, and the 2 Nov run $89,508.80 (Brookvale, Contact Group, Spacera, Easton Belle, Hire Intelligence, Type Transcripts, Ricoh).',
 'O52':'Q1 BAS did not go out on 28 Oct. It is $115,000 and is in the November forecast on 25 Nov through the agent.',
 'O59':'Cheque statement balance plus savings. Savings is $410,000 at 28 Sep moved by the transfers on the statement and the interest: $250,000 at 7 Oct, $550,000 at 14 Oct, $400,000 at 21 and 28 Oct, $400,411.39 at 1 Nov. Week 5 cheque is worked forward from the 1 Nov statement with the 2 to 4 Nov items and the $150,000 of transfers - replace with the CommBiz balance at 4 Nov.',
 'O60':'Nil for weeks 1 to 4 - the actuals tie to the cheque statement. Week 5 will tie once the real 4 Nov balance is in.',
}
for k,v in notes.items(): cf[k]=v
cf['A69']="1. Aware Super paid the $237,987 on 13 Oct and everything else by 27 Oct - risk cleared. The whole Complete AV bill ($178,119) went out ahead of it."
cf['A70']="2. Complete AV labour and programming ($39,033 inc GST) was released and paid 8 Oct, so INV41565 is fully paid. The $5,334 balance of the quote is still to be invoiced (claim 3)."
cf['A71']="3. Still overdue at 4 Nov: Downer $39,664 (both from August), Microsoft $13,846, Ventia $10,016, ASX $5,748, ON24 $3,006 and a few small new customers - about $85k. Ventia paid $26,661 on 2 Nov and Gen-e cleared, so they are moving. Chase Downer and Microsoft this week."
cf['A72']="4. BAS - Sep quarter is $115,000 and did not go out on 28 Oct. It is in the November forecast on 25 Nov through the agent. The other states payroll tax ($2,684) for September is not on the cheque statement - check it was paid."
# ================= RECEIVABLES to Xero at 4 Nov =================
for r in range(5,66):
    for c in range(1,12):
        cell=rec.cell(r,c); cell.value=None; cell.fill=NOFILL
base=[copy.copy(rec.cell(36,c)._style) for c in range(1,12)]
r=5
for row in inv:
    for c in range(1,12): rec.cell(r,c)._style=copy.copy(base[c-1])
    for c,v in zip('ABCDEFGH',row[:8]): rec[f'{c}{r}']=v
    rec[f'I{r}']=f'=H{r}-E{r}'
    rec[f'J{r}']=f'=IF(H{r}<Assumptions!$B$5,"Before",IF(H{r}>Assumptions!$C$48,"After",INT((H{r}-Assumptions!$B$5)/7)+1))'
    rec[f'K{r}']=row[8]; r+=1
rec['A2']=("Xero aged receivables at 4 Nov ($726,549.38, 56 invoices - the small CBA ones are grouped). The October contract invoices went out 30 Oct. Expected date (blue) is my call based on how each customer has been paying. "
           "This list feeds the November forecast file; the forecast columns on the Cash Flow tab here are the 28 Sep forecast, frozen.")
rec['A67']='of which open in Xero at 4 Nov (should be $726,549.38)'; rec['F67']='=SUMIFS(F5:F65,G5:G65,"<>To be raised")'
rec['A68']='of which expected inside this forecast (28 Sep to 4 Nov) - nil, these are all still to come'
rec['A69']='of which expected after 4 Nov'
# ================= PAYABLES to Xero at 4 Nov =================
tmpl=[copy.copy(pb.cell(60,c)._style) for c in range(1,16)]
for r in range(19,76):
    for c in range(1,16):
        cell=pb.cell(r,c); cell.value=None; cell._style=copy.copy(tmpl[c-1])
    pb[f'H{r}']=f'=IF(G{r}<Assumptions!$B$5,"Before",IF(G{r}>Assumptions!$C$48,"After",INT((G{r}-Assumptions!$B$5)/7)+1))'
    pb[f'K{r}']=f'=IF(AND(\'Creditors Run\'!$B$5=1,I{r}="Xero bill",LEFT(J{r},4)<>"PAID"),"No - in the creditors run","Yes")'
def nbd(d): return d if d.weekday()<5 else d+datetime.timedelta(days=7-d.weekday())
r=19
for b in bills:
    sup,ref,due,amt,meth,cat,note,net,gst,invd,payon=b
    for c,v in zip('ABCDEFGIJ',(sup,ref,due,amt,meth,cat,payon,'Xero bill',note)): pb[f'{c}{r}']=v
    r+=1
for t in templates:
    sup,ref,due,amt,meth,cat,src,note=t
    for c,v in zip('ABCDEFGIJ',(sup,ref,due,amt,meth,cat,nbd(due),src,note)): pb[f'{c}{r}']=v
    r+=1
pb['A77']='TOTAL bills listed from Xero (aged payables at 4 Nov, should be $84,568.84)'
pb['A2']=("Xero aged payables at 4 Nov ($84,568.84, 29 bills) plus the repeating bill templates for November. The 2 Nov creditors run cleared Brookvale, Contact Group, Spacera, Easton Belle and the small ones. "
          "Card rows get paid on the PL/DL cards and sit in the credit card line. This list feeds the November forecast file.")
pb['A4']='PAYMENTS THIS WEEK (5 - 11 Nov)'
block=[
 ('=A19','=B19','=C19','=D19','=E19','AGREED - PAY','=G19','Appspace - overdue, Aware paid us for it 20 Oct. Pay Fri 6 Nov.'),
 ('=A20','=B20','=C20','=D20','=E20','AGREED - PAY','=G20','Start Webcast - DL approved. Pay Fri 6 Nov.'),
 ('Headroom Audio Visual','INV-0034 + INV15 + INV-0033 (check INV15 / INV-0033 are not the same job)',D(2026,11,6),'=SUM(D21:D23)','EFT','Due this week',D(2026,11,6),'MH approved. $18,723 if all three are real.'),
 ('2020 Events / WordBirds / Dubology / Hire Intelligence / Deaf Services','Event hire bills due 4 to 10 Nov',D(2026,11,10),'=SUM(D24:D29)+D39','EFT','Due this week',D(2026,11,9),'All MH approved. Deaf Services needs DL / TW.'),
 ('=A40','=B40','=C40','=D40','=E40','Due this week','=G40','Employsure Oct - a month overdue.'),
 ('Optus / Posh / Telstra / Officeworks','Small ones due by 9 Nov',D(2026,11,9),'=D41+D42+D44+D45','BPAY / EFT','Due this week',D(2026,11,9),'Routine.'),
 ('=A48','=B48','=C48','=D48','=E48','Due this week','=G48','Employsure Nov - template.'),
 ('=A49','=B49','=C49','=D49','=E49','Due this week','=G49','8020 Advisors - template, check the invoice has come in.'),
 ('Payroll tax','October return - NSW + other states',D(2026,11,7),8602.75,'EFT','Due this week',D(2026,11,6),'Due Sat 7 Nov, pay Fri 6 Nov. Not a bill - on the Cash Flow employment line.'),
]
green=[copy.copy(pb.cell(6,c)._style) for c in range(1,16)]; plain=[copy.copy(pb.cell(7,c)._style) for c in range(1,16)]
for i,row in enumerate(block):
    rr=6+i
    for c in range(1,16):
        pb.cell(rr,c)._style=copy.copy(green[c-1] if row[5]=='AGREED - PAY' else plain[c-1])
        if row[5]!='AGREED - PAY': pb.cell(rr,c).fill=NOFILL
    for c,v in zip('ABCDEFGJ',row): pb[f'{c}{rr}']=v
pb['A16']='of which agreed / due this week'
# ================= CREDITORS RUN to Xero at 4 Nov =================
for r in range(11,511):
    for c in range(1,14): cr.cell(r,c).value=None
cr['B4']=D(2026,12,1); cr['C4']='Tue 1 Dec - the next monthly run. Anything with its own date in column L ignores this, and everything open at 4 Nov is dated to its due date.'
cr['B7']=84568.84
for i,b in enumerate(bills):
    sup,ref,due,amt,meth,cat,note,net,gst,invd,payon=b; rr=11+i
    code,_,desc=ref.partition(' - ')
    cr[f'A{rr}']=invd; cr[f'B{rr}']='Bill'; cr[f'C{rr}']=code; cr[f'E{rr}']=f'{sup} - {desc}'; cr[f'G{rr}']=net; cr[f'H{rr}']=gst; cr[f'I{rr}']=amt; cr[f'J{rr}']=amt; cr[f'K{rr}']='Awaiting Payment'
    if meth=='Card': cr[f'M{rr}']='x'
    else: cr[f'L{rr}']=payon
for rr in range(40,60): cr[f'A{rr}']=None
cr['A41']='Rows 11 to 39 are the 29 open bills in Xero at 4 Nov ($84,568.84) typed in the report layout. Next month, clear them and paste the new report over the top.'
cr['A41'].font=Font(italic=True,size=9)
a['C6']='Cheque $25,155 + savings $410,000 at 28 Sep, both per CommBiz. Not counting the other accounts.'
wb.calculation.fullCalcOnLoad=True
wb.save('out/oct_bank.xlsx'); print('saved')
