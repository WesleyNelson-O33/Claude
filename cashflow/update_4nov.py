import openpyxl, datetime, copy, re
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
D=datetime.datetime
wb=openpyxl.load_workbook('out/roll.xlsx')
cf=wb['Cash Flow']; a=wb['Assumptions']; p=wb['Payroll']; rec=wb['Receivables']; pb=wb['Payables']; cr=wb['Creditors Run']
NOFILL=PatternFill(fill_type=None); GAP=456746.27
def nbd(d):
    return d if d.weekday()<5 else d+datetime.timedelta(days=7-d.weekday())

# ================= WEEK 1 ACTUALS (29 Oct - 4 Nov) =================
act={13:'=ROUND(59581.86*1.1,2)',14:411.39,19:'=843985.1-752991.57',20:0,21:0,22:0,23:0,
     27:'=ROUND(1227.5*1.1,2)',28:'=ROUND(8212.5*1.1,2)',29:0,30:0,31:'=ROUND((300+156.25)*1.1,2)',
     35:'=ROUND((20681.16+4894.76+692.91)*1.1,2)',36:0,37:0,38:0,39:0,40:0,41:0,42:0,43:0,44:0,45:'=ROUND(6.92*1.1,2)+10',46:0,47:0,
     48:'=66238.98+10555.92+5728.8+3932.5+2585+257.4+210.2',52:0,59:f'=752991.57-{GAP}'}
for r,v in act.items():
    cf[f'C{r}']=v; cf[f'C{r}']._style=copy.copy(cf['C13']._style)
cf['C8']=1
cf['O8']='Week 1 is done (flag 1) off Xero to Wed 4 Nov. Weeks 2 to 5 are forecast. Set each flag to 1 as the week closes and the CommBiz balance goes in.'
cf['O13']='Week 1 actual: PwC $20,963, Ventia $26,661 (two of the three), CBA foldback $3,498, YSG, Garvan, Elkiem. Week 2: Deloitte Oct $117k, Bankwest, CBA Brisbane, the overdue small ones. Week 3: CBA production kit $118k plus the CBA Oct contract invoices. Week 5: ASX AGM $70.6k, TLC $39k, April 5 $32k, PwC, Aware.'
cf['O19']='Payroll tab - gross x net %. Tue 3 Nov went out at $90,994 net ($115,121 gross per Xero). Next: Tue 17 Nov (week 3) and Tue 1 Dec (week 5).'
cf['O20']='Gross x 12%. The 3 Nov super ($14,631) had not left the bank at 4 Nov - it lands in week 2.'
cf['O42']='Brookvale $66,239 went out in the week 1 creditors run. Nothing else expected.'
cf['O48']='Creditors Run tab. Week 1 actual is what went out in the Mon 2 Nov run - Brookvale $66k, Contact Group $10.6k, Spacera, Easton Belle, Hire Intelligence, Type Transcripts, Ricoh ($89.5k). Appspace $19k and Start Webcast $12.8k did not go - now in week 2. Next run Tue 1 Dec.'
cf['O59']='Type the real closing balance (cheque + savings) in at the end of each week. Week 1 is worked back from Xero total cash at 4 Nov ($752,992) less the other accounts ($456,746) - replace with the CommBiz figure.'

# ================= ASSUMPTIONS / PAYROLL =================
a['B9']=113400; a['C9']='Average of the last two runs: 20 Oct $111,704 and 3 Nov $115,121 (Xero cash P&L). Detail on the Payroll tab.'
a['B10']=0.79; a['C10']='3 Nov run: $115,121 gross, $90,994 left the bank = 79%. Swap in the ABA figure when you have it.'
a['C21']="It did not go out on 28 Oct, so I've assumed it goes through the agent and is paid 25 Nov (week 4). Change the date in B25 if it goes earlier, or set this to 0."
p['J5']='Paid Tue 3 Nov - $90,994 net left the bank (Xero), $115,121 gross, super $14,631 to follow. Done.'
p['A24']='Average gross per fortnight (last two runs)'; p['B24']='=(C23+115121.14)/2'; p['C24']='20 Oct $111,704 and 3 Nov $115,121 (Xero cash P&L 29 Oct - 4 Nov).'

# ================= RECEIVABLES =================
for r in range(5,66):
    for c in range(1,12):
        cell=rec.cell(r,c); cell.value=None; cell.fill=NOFILL
base=[copy.copy(rec.cell(5,c)._style) for c in range(1,12)]
def rrow(r,vals,note):
    for c in range(1,12): rec.cell(r,c)._style=copy.copy(base[c-1])
    for c,v in zip('ABCDEFGH',vals): rec[f'{c}{r}']=v
    rec[f'I{r}']=f'=H{r}-E{r}'
    rec[f'J{r}']=f'=IF(H{r}<Assumptions!$B$5,"Before",IF(H{r}>Assumptions!$C$48,"After",INT((H{r}-Assumptions!$B$5)/7)+1))'
    rec[f'K{r}']=note
inv=[
 ('Deloitte Services Trust','INV-10687','Oct contract (CONTRACT)',D(2026,10,30),D(2026,11,29),116099.99,'Current',D(2026,11,9),'Deloitte paid the Sep contract invoice 8 days after we sent it. Mon 9 Nov.'),
 ('Deloitte Services Trust','INV-10702','DTTL additionals October 2026',D(2026,10,30),D(2026,11,29),1186.57,'Current',D(2026,11,9),'With the contract one.'),
 ('Commonwealth Bank (CBA Place South)','INV-10677','Bankwest production kit (PO3311640)',D(2026,10,16),D(2026,11,15),118250.0,'Current',D(2026,11,16),'Biggest one in the forecast. Not in at 4 Nov, so assume the due date - Sun 15 Nov, call it Mon 16 Nov (week 3).'),
 ('Commonwealth Bank (CBA Place South)','INV-10686','Oct contract (85720001CBPSCSJUN27)',D(2026,10,30),D(2026,11,29),24871.11,'Current',D(2026,11,16),'Sep ones were paid 15 days after invoice - Mon 16 Nov.'),
 ('Commonwealth Bank (CBA Place South)','INV-10712 / 10714 / 10717 / 10718 / 10719','Oct after hours + dedicated tech (5 small invoices)',D(2026,10,30),D(2026,11,29),3828.0,'Current',D(2026,11,16),'$2,145 + $467.50 x 2 + $374 x 2. With the contract one.'),
 ('Commonwealth Bank (CBA Place South)','INV-10727','Live event support October 2026',D(2026,11,4),D(2026,12,4),11024.75,'Current',D(2026,11,19),'15 days after invoice - Thu 19 Nov.'),
 ('Commonwealth Bank (CBA Square)','INV-10691','Oct contract (85720001CBSCSJUN27)',D(2026,10,30),D(2026,11,29),11792.0,'Current',D(2026,11,16),'Same as Place South.'),
 ('Commonwealth Bank (CBA Square)','INV-10722','Oct additional (85720001CBSCSJUN27)',D(2026,10,30),D(2026,11,29),2592.7,'Current',D(2026,11,16),'Same.'),
 ('Commonwealth Bank (CBA Melbourne)','INV-10697','Oct contract (85720001MEL435CSJUN27)',D(2026,10,30),D(2026,11,29),13156.0,'Current',D(2026,11,16),'Same.'),
 ('Commonwealth Bank (CBA Melbourne)','INV-10711 / 10713 / 10715 / 10716','Oct after hours + dedicated tech (4 small invoices)',D(2026,10,30),D(2026,11,29),2618.0,'Current',D(2026,11,16),'$748 x 3 + $374. Same.'),
 ('Commonwealth Bank (CBA Brisbane)','INV-10685','Oct contract (85720001BRISCSJUN27)',D(2026,10,30),D(2026,11,29),8177.95,'Current',D(2026,11,9),'Sep one was paid a week after invoice - Mon 9 Nov.'),
 ('Bankwest','INV-10690','Oct contract (85720001BWPCSJUN27)',D(2026,10,30),D(2026,11,29),13156.0,'Current',D(2026,11,6),'Bankwest paid the Sep one two days after we sent it. Not in at 4 Nov - Fri 6 Nov.'),
 ('Bankwest','INV-10688','Oct additionals',D(2026,10,30),D(2026,11,29),1584.0,'Current',D(2026,11,6),'With the contract one.'),
 ('PwC Services','INV-10701','PO2601702104 - Oct contract',D(2026,10,30),D(2026,11,29),13097.7,'Current',D(2026,11,30),'PwC pay on the due date (Sep ones came 29 Oct). Due Sun 29 Nov so Mon 30 Nov.'),
 ('PwC Services','INV-10689','PO2601702371 - Oct contract',D(2026,10,30),D(2026,11,29),7865.0,'Current',D(2026,11,30),'Same.'),
 ('PwC Services','INV-10698','PO2601704199 - AV equipment',D(2026,10,30),D(2026,11,29),7201.7,'Current',D(2026,11,30),'Same.'),
 ('Aware Super','INV-10699','Melbourne display replacements - equipment',D(2026,10,30),D(2026,11,29),16493.4,'Current',D(2026,11,27),'Aware paid the Sep maintenance three days before due. Fri 27 Nov.'),
 ('Aware Super','INV-10700','Annual National Maintenance - October',D(2026,10,30),D(2026,11,29),12100.0,'Current',D(2026,11,27),'Same.'),
 ('Aware Super','INV-10720','Occupancy sensor upgrade - engineering',D(2026,10,30),D(2026,11,29),2977.7,'Current',D(2026,11,27),'Same.'),
 ('Cochlear','INV-10725','AGM 26.10.2026',D(2026,11,4),D(2026,12,4),70302.1,'Current',D(2026,12,3),'Cochlear paid the Aug results webcast a day before due - Thu 3 Dec, just outside the forecast (ends Wed 2 Dec). Upside if it comes early.'),
 ('ASX Operations','INV-10668','External listing test days 29.9.26',D(2026,9,30),D(2026,10,30),5747.5,'Overdue',D(2026,11,6),'A week overdue - ASX normally pay on the day. Chase. Fri 6 Nov.'),
 ('ASX Operations','INV-10705','ASX AGM 22.10.2026',D(2026,10,29),D(2026,11,28),70644.76,'Current',D(2026,11,30),'Due Sat 28 Nov, call it Mon 30 Nov. The big one for ASX.'),
 ('ASX Operations','INV-10704','CEO Connect series 27.10.2026',D(2026,10,29),D(2026,11,28),4041.4,'Current',D(2026,11,30),'With the AGM one.'),
 ('ASX Operations','INV-10703','Listing event lighting 14.10.26',D(2026,10,29),D(2026,11,28),1903.0,'Current',D(2026,11,30),'With the AGM one.'),
 ('April 5','INV-10706','Microsoft SME&C event 14.10.2026 (PO016066)',D(2026,10,29),D(2026,11,28),32452.77,'Current',D(2026,11,30),'New customer - on the due date, Mon 30 Nov.'),
 ('The Lottery Corporation','INV-10710','AGM 14.10.2026 (variation)',D(2026,10,29),D(2026,11,28),39187.49,'Current',D(2026,11,30),'TLC paid the Aug results three days before due. Mon 30 Nov.'),
 ('The Lottery Corporation','INV-10682','AGM event photography 14.10.2026',D(2026,10,15),D(2026,11,14),1534.5,'Current',D(2026,11,16),'Due Sat 14 Nov so Mon 16 Nov.'),
 ('MUFG Corporate Markets','INV-10709','IMB Limited AGM 27.10.2026',D(2026,10,29),D(2026,11,28),9129.73,'Current',D(2026,11,30),'MUFG paid the Aug ones on the due date. Mon 30 Nov.'),
 ('Downer EDI','INV-10611','Full year employee briefing 25.08.2026',D(2026,8,31),D(2026,9,30),26855.13,'Overdue 1 month',D(2026,11,13),'Five weeks overdue and nothing in. Chase. Fri 13 Nov.'),
 ('Downer EDI','INV-10587','Full year results briefing 20.08.2026',D(2026,8,25),D(2026,9,24),12808.69,'Overdue 1 month',D(2026,11,13),'Six weeks overdue. Chase with the other one.'),
 ('Ventia','INV-10632','RUOK Day event 10.09.2026',D(2026,9,16),D(2026,10,16),10015.89,'Overdue',D(2026,11,13),'Ventia paid the two Aug ones on 2 Nov, so this one should follow. Fri 13 Nov.'),
 ('Microsoft','INV-10554','PO 10160645',D(2026,7,31),D(2026,8,30),13845.7,'Overdue 2 months',D(2026,11,20),'Two months overdue - usually a PO / portal issue. Chase. Week 4.'),
 ('ON24','INV-10589','PO25962',D(2026,8,26),D(2026,9,25),3005.75,'Overdue 1 month',D(2026,11,20),'A month overdue - chase.'),
 ('Chorus Call Australia','INV-10678','Redox AGM 07.10.2026',D(2026,10,16),D(2026,11,15),5326.2,'Current',D(2026,11,16),'Due Sun 15 Nov so Mon 16 Nov.'),
 ("St Luke's Liverpool",'INV-10676','Digital mixer equipment',D(2026,9,30),D(2026,10,30),4988.5,'Overdue',D(2026,11,6),'New customer, a few days overdue - chase. Fri 6 Nov.'),
 ('Global Philanthropic','INV-10670','Talking Philanthropy Aus 18.09.2026',D(2026,9,30),D(2026,10,30),4343.9,'Overdue',D(2026,11,6),'New customer, a few days overdue - chase.'),
 ('Global Philanthropic','INV-10669','Talking Philanthropy event B-roll',D(2026,9,30),D(2026,10,30),1373.9,'Overdue',D(2026,11,6),'With the other one.'),
 ('ICC Sydney','INV-10672','Speaker presentation technicians Sep (TTS)',D(2026,9,30),D(2026,10,30),4100.26,'Overdue',D(2026,11,6),'ICC paid the other Sep one early - this one is a few days late. Chase.'),
 ('Department of Agriculture, Fisheries and Forestry','INV-10721','Room standards drafting',D(2026,10,30),D(2026,11,29),3850.0,'Current',D(2026,11,30),'Government - on the due date, Mon 30 Nov.'),
 ('Datacom Systems','INV-10723','DFATHP Poly phone modifications (PO417675)',D(2026,10,30),D(2026,11,29),1262.8,'Current',D(2026,11,30),'On the due date.'),
 ('Vega Global Australia','INV-10647','Payments & Securities Forums 14-17.9.26',D(2026,9,22),D(2026,10,22),1969.0,'Overdue',D(2026,11,6),'Two weeks overdue - chase.'),
 ('Vega Global Australia','INV-10684','AI Forum audio assistance 13-14.10.26',D(2026,10,16),D(2026,11,15),1907.13,'Current',D(2026,11,16),'Due Sun 15 Nov so Mon 16 Nov.'),
 ('Garvan Research Foundation','INV-10680','Breast cancer symposium (PO-59639)',D(2026,10,15),D(2026,11,14),836.0,'Current',D(2026,11,16),'Garvan paid the Suttons one 10 days after invoice. Mon 16 Nov.'),
 ('Garvan Research Foundation','INV-10679','Breast cancer public lecture 01.10.2026',D(2026,10,15),D(2026,11,14),418.0,'Current',D(2026,11,16),'Same.'),
 ('APT Management Services (APA)','INV-10726','AV train the trainer 21.10.26 (PO-00034822)',D(2026,11,4),D(2026,12,4),863.5,'Current',D(2026,11,9),'APA pay within days of the invoice.'),
 ('Dimensional Fund Advisors','INV-10692','AV audio support 15.10.26',D(2026,10,20),D(2026,11,19),418.0,'Current',D(2026,11,6),'Paid the last one 9 days after invoice - already past that. Fri 6 Nov.'),
 ('Noisebox Integrated Technologies','INV-10681','Audio dry hire 13.10.16',D(2026,10,16),D(2026,11,15),203.5,'Current',D(2026,11,16),'Due Sun 15 Nov so Mon 16 Nov.'),
 ('Jaazaniah Salanoa (staff)','INV-10621','Goget hire recovery',D(2026,8,31),D(2026,9,30),66.64,'Overdue',D(2026,12,4),'Staff recovery - goes through payroll, no cash. Left out.'),
 ('Meet Magic','INV-10464','Human AI event highlights videos 29.05.2026',D(2026,6,19),D(2026,6,26),5075.07,'Overdue 4 months',D(2026,12,18),'DOUBTFUL - four months overdue. Left out. Needs escalating.'),
]
assert len(inv)==49
tot=sum(x[5] for x in inv); assert abs(tot-726549.38)<0.02, tot
r=5
for row in inv: rrow(r,row[:8],row[8]); r+=1
rec['A2']=("Xero aged receivables at 4 Nov ($726,549.38, 56 invoices - the small CBA ones are grouped). The October contract invoices went out 30 Oct. Expected date (blue) is my call based on how each customer has been paying - change it and the forecast moves. "
           "Gen-e, Ventia (two of three), PwC and YSG all paid in the last week.")
rec['A67']='of which open in Xero at 4 Nov (should be $726,549.38)'
rec['A69']='of which pushed past the forecast (Cochlear $70k due 4 Dec, Meet Magic, Jaaz)'

# ================= PAYABLES =================
tmpl=[copy.copy(pb.cell(60,c)._style) for c in range(1,16)]
for r in range(19,96):
    for c in range(1,16):
        cell=pb.cell(r,c); cell.value=None; cell._style=copy.copy(tmpl[c-1])
    pb[f'H{r}']=f'=IF(G{r}<Assumptions!$B$5,"Before",IF(G{r}>Assumptions!$C$48,"After",INT((G{r}-Assumptions!$B$5)/7)+1))'
    pb[f'K{r}']=f'=IF(AND(\'Creditors Run\'!$B$5=1,I{r}="Xero bill",LEFT(J{r},4)<>"PAID"),"No - in the creditors run","Yes")'
RUN=D(2026,12,1)
bills=[
 ('Appspace','INV00133660 - Appspace licensing for Aware Super (balance)',D(2026,10,24),19077.66,'EFT','Client subscriptions','Overdue - did not go in the 2 Nov run. Aware paid us for it 20 Oct. Pay Fri 6 Nov.',17343.33,1734.33,D(2026,9,24),D(2026,11,6)),
 ('Start Webcast','INV-202222 - CBA AGM 2026 equipment & services',D(2026,10,29),12787.06,'EFT','Equipment & service hires','DL has approved it now. Pay Fri 6 Nov.',11624.6,1162.46,D(2026,10,14),D(2026,11,6)),
 ('Headroom Audio Visual','INV-0034 - Ventia (MELB) CEO town hall equipment hire',D(2026,11,6),12739.1,'EFT','Equipment & service hires','MH approved. On the due date.',11581.0,1158.1,D(2026,11,4),D(2026,11,6)),
 ('Headroom Audio Visual','INV15 - CBA foldback monitor dry hire 21.10.26',D(2026,10,30),2992.0,'EFT','Equipment & service hires','MH approved. Overdue - pay Fri 6 Nov. Check it is not the same job as INV-0033.',2720.0,272.0,D(2026,10,25),D(2026,11,6)),
 ('Headroom Audio Visual','INV-0033 - InnovaTech / CBA foldback monitor dry hire 21.10.26',D(2026,11,6),2992.0,'EFT','Equipment & service hires','MH approved. Same amount and job as INV15 - check before paying both.',2720.0,272.0,D(2026,11,4),D(2026,11,6)),
 ('2020 Events','INV-1611 - Microsoft SME&C townhall 14.10.26 hire',D(2026,11,4),2346.37,'EFT','Equipment & service hires','MH approved. Due 4 Nov - pay Fri 6 Nov.',2133.06,213.31,D(2026,10,21),D(2026,11,6)),
 ('2020 Events','INV-1618 - Cochlear AGM 26.10.26 hire',D(2026,11,9),5665.0,'EFT','Equipment & service hires','MH approved, needs the quote to split costs. On the due date.',5150.0,515.0,D(2026,10,26),D(2026,11,9)),
 ('WordBirds','INV-2016 - ASX AGM 22.10.26 services',D(2026,11,9),5063.5,'EFT','Equipment & service hires','MH approved. On the due date.',4603.18,460.32,D(2026,10,26),D(2026,11,9)),
 ('WordBirds','INV-2019 - Cochlear AGM 26.10.26 services',D(2026,11,10),3267.0,'EFT','Equipment & service hires','MH approved. On the due date.',2970.0,297.0,D(2026,10,27),D(2026,11,10)),
 ('Dubology Studios','INV26-12398 - Microsoft SME&C and CBA LES Oct services',D(2026,11,9),2117.5,'EFT','Equipment & service hires','MH approved. On the due date.',1925.0,192.5,D(2026,10,26),D(2026,11,9)),
 ('Hire Intelligence','CS167742 - Chorus Call Nanosonics AGM 4.11.26 equipment hire',D(2026,11,3),2002.0,'EFT','Equipment & service hires','MH approved. Due 3 Nov - pay Fri 6 Nov.',1820.0,182.0,D(2026,11,3),D(2026,11,6)),
 ('Type Transcripts','INV6065 - ASX AGM and Cochlear AGM transcription',D(2026,11,13),1378.52,'EFT','Equipment & service hires','MH approved. On the due date.',1253.2,125.32,D(2026,10,29),D(2026,11,13)),
 ('Riedel Communications','INV00010410 - Microsoft SME&C townhall 14.10.26 hire',D(2026,11,14),885.06,'EFT','Equipment & service hires','MH approved. Due Sat 14 Nov so Mon 16 Nov.',804.6,80.46,D(2026,10,15),D(2026,11,16)),
 ('Riedel Communications','INV00010482 - MUFG Suncorp investor day 29.10.2026 hire',D(2026,11,25),1829.52,'EFT','Equipment & service hires','MH approved. On the due date.',1663.2,166.32,D(2026,10,26),D(2026,11,25)),
 ('Riedel Communications','INV00010469 - ASX AGM 22.10.26 hire',D(2026,11,26),1834.27,'EFT','Equipment & service hires','MH approved. On the due date.',1667.52,166.75,D(2026,10,27),D(2026,11,26)),
 ('Chorus Call Australia','INV02-82515 - Cochlear AGM 26.10.26 services',D(2026,11,27),2065.84,'EFT','Equipment & service hires','MH approved. On the due date.',1878.04,187.8,D(2026,10,28),D(2026,11,27)),
 ('Access Innovation Media','INV22273 - CBA AGM 2026 services',D(2026,11,18),924.0,'EFT','Equipment & service hires','DL / TW to approve.',840.0,84.0,D(2026,10,19),D(2026,11,18)),
 ('Access Innovation Media','INV22274 - CBA AGM 2026 services',D(2026,11,18),682.0,'EFT','Equipment & service hires','DL / TW to approve.',620.0,62.0,D(2026,10,19),D(2026,11,18)),
 ('Access Innovation Media','INV22272 - CBA AGM 2026 services',D(2026,11,18),264.0,'EFT','Equipment & service hires','DL / TW to approve.',240.0,24.0,D(2026,10,19),D(2026,11,18)),
 ('Access Innovation Media','INV22275 - CBA AGM 2026 services',D(2026,11,18),165.0,'EFT','Equipment & service hires','DL / TW to approve.',150.0,15.0,D(2026,10,19),D(2026,11,18)),
 ('Deaf Services','INV409016 - CBA AGM 2026 services',D(2026,11,9),310.75,'EFT','Equipment & service hires','DL / TW to approve and confirm the cost centre.',282.5,28.25,D(2026,10,26),D(2026,11,9)),
 ('Employsure','42 of 60 - October',D(2026,10,5),880.0,'EFT','Professional fees','A month overdue. Pay Fri 6 Nov.',800.0,80.0,D(2026,10,5),D(2026,11,6)),
 ('Optus','INV000600307641 - staff mobiles Oct',D(2026,11,9),840.98,'BPAY','Telephone & internet','BPAY on the due date.',764.53,76.41,D(2026,10,10),D(2026,11,9)),
 ('Posh Services (Ambius)','INV21742621 - indoor plants Oct',D(2026,11,8),245.0,'EFT','Office','Due Sun 8 Nov so Mon 9 Nov.',222.73,22.27,D(2026,10,9),D(2026,11,9)),
 ('Ricoh Australia','INV15788514 - printing Oct',D(2026,11,30),257.47,'DD','Office','Direct debit.',234.06,23.41,D(2026,10,26),D(2026,11,30)),
 ('Telstra','TTI20233644300 - Nighthawk Oct',D(2026,11,6),136.0,'BPAY','Telephone & internet','BPAY on the due date.',123.64,12.36,D(2026,10,25),D(2026,11,6)),
 ('Officeworks','INV632570641 - keyboard & mouse for CBA AGM',D(2026,11,6),108.0,'EFT','Office','DL approved. On the due date.',98.18,9.82,D(2026,10,7),D(2026,11,6)),
 ('Energy Australia','INV260412412804 - electricity 23 Aug to 21 Sep',D(2026,10,14),696.25,'Card','Electricity','DD via the PL card.',632.95,63.3,D(2026,9,24),D(2026,10,14)),
 ('Google','INV1115.7 - YouTube Premium Oct',D(2026,10,28),16.99,'Card','Dues & subscriptions','DL card.',15.45,1.54,D(2026,10,28),D(2026,10,28)),
]
assert len(bills)==29 and abs(sum(b[3] for b in bills)-84568.84)<0.02, sum(b[3] for b in bills)
templates=[
 ('Employsure','43 of 60 - November',D(2026,11,5),880.0,'EFT','Professional fees','Repeating template','Template, due 5 Nov.'),
 ('Cleveland Partners (8020 Advisors)','Monthly advisory fees - November',D(2026,11,9),3850.0,'EFT','Professional fees','Repeating template','Template $3,850 a month. Nothing paid in Oct that I can see in Xero - check.'),
 ('Telstra','Recharges x 4 - November',D(2026,11,9),176.0,'Card','Telephone & internet','Repeating template','PL card.'),
 ('LiveU','SOLO standard data pack - November',D(2026,11,14),1485.0,'EFT','Client subscriptions','Repeating template','Due Sat 14 Nov so Mon 16 Nov.'),
 ('First Focus IT','CORE Intelligent Managed Services - November',D(2026,11,16),13755.77,'DD','IT support','Repeating template','DD on the due date.'),
 ('First Focus IT','Managed Services Azure - November',D(2026,11,16),710.5,'DD','IT support','Repeating template','DD on the due date.'),
 ('Insphire (Current RMS)','November',D(2026,11,16),431.0,'Card','Dues & subscriptions','Repeating template','PL card.'),
 ('Mailchimp','November',D(2026,11,18),37.62,'Card','Dues & subscriptions','Repeating template','PL card.'),
 ('Optus','Lugstein home internet - November',D(2026,11,20),101.2,'BPAY','Telephone & internet','Repeating template','BPAY on the due date.'),
 ('RJW Group Holdings','Digital growth marketing - November',D(2026,11,21),4675.0,'EFT','Advertising','Repeating template','Due Sat 21 Nov so Mon 23 Nov.'),
 ('Telstra','Nighthawk data plans - November',D(2026,11,21),130.0,'BPAY','Telephone & internet','Repeating template','Due Sat 21 Nov so Mon 23 Nov.'),
 ('Optus','Office internet & landline - November',D(2026,11,21),629.66,'Card','Telephone & internet','Repeating template','PL card.'),
 ('Xero / Adobe / Spotify / Avangate / misc','Small monthly subscriptions',D(2026,11,24),700.0,'Card','Dues & subscriptions','Repeating template','PL / DL cards - all in the card line.'),
 ('Zoom','Monthly subscription - November',D(2026,11,28),323.97,'Card','Dues & subscriptions','Repeating template','PL card.'),
 ('Employment Hero','Subscription - October',D(2026,11,29),2048.2,'Card','Dues & subscriptions','Repeating template','PL card - in the card line.'),
 ('Goget Carshare','October usage',D(2026,11,29),1783.07,'Card','Direct travel & freight','Repeating template','PL card - in the card line.'),
 ('SPACERA','Aware Super national support - November',D(2026,12,1),5728.8,'EFT','Client subscriptions','Repeating template','Due Tue 1 Dec - week 5, with the Dec run.'),
 ('Investa Asset Management','Rent December 2026',D(2026,12,1),28895.72,'EFT','Rent','Repeating template','Billed 1 Dec. Nov rent was paid 2 Nov with the run, so Dec lands Tue 1 Dec (week 5).'),
 ('Posh Services (Ambius)','Indoor plants - November',D(2026,12,9),245.0,'EFT','Office','Repeating template','After this forecast.'),
 ('Optus','Staff mobiles - November',D(2026,12,9),804.06,'BPAY','Telephone & internet','Repeating template','After this forecast.'),
 ('The CEO Institute','Executive membership (JS) - quarterly',D(2026,12,7),2310.0,'EFT','Dues & subscriptions','Repeating template','After this forecast.'),
 ('AVIXA','DL annual silver membership',D(2026,12,11),2850.0,'EFT','Dues & subscriptions','Repeating template','After this forecast.'),
]
r=19
for b in bills:
    sup,ref,due,amt,meth,cat,note,net,gst,invd,payon=b
    for c,v in zip('ABCDEFGIJ',(sup,ref,due,amt,meth,cat,payon,'Xero bill',note)): pb[f'{c}{r}']=v
    r+=1
for t in templates:
    sup,ref,due,amt,meth,cat,src,note=t
    for c,v in zip('ABCDEFGIJ',(sup,ref,due,amt,meth,cat,nbd(due),src,note)): pb[f'{c}{r}']=v
    r+=1
pb['A97']='TOTAL bills listed from Xero (aged payables at 4 Nov, should be $84,568.84)'
pb['A2']=("Xero aged payables at 4 Nov ($84,568.84, 29 bills) plus the repeating bill templates for November. The 2 Nov creditors run cleared Brookvale, Contact Group, Spacera, Easton Belle and the small ones. "
          "Card rows get paid on the PL/DL cards and sit in the credit card line on the Cash Flow tab. With the Creditors Run tab switched on, the open Xero bills come through the Creditors run row and the templates come through the category lines.")
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
 ('Payroll tax (Assumptions tab)','October return - NSW + other states',D(2026,11,7),'=Assumptions!B17+Assumptions!B18','EFT','Due this week',D(2026,11,6),'Due Sat 7 Nov, pay Fri 6 Nov. Not a bill - on the Cash Flow employment line.'),
]
green=[copy.copy(pb.cell(6,c)._style) for c in range(1,16)]
plain=[copy.copy(pb.cell(7,c)._style) for c in range(1,16)]
for i,row in enumerate(block):
    rr=6+i
    for c in range(1,16):
        pb.cell(rr,c)._style=copy.copy(green[c-1] if row[5]=='AGREED - PAY' else plain[c-1])
        if row[5]!='AGREED - PAY': pb.cell(rr,c).fill=NOFILL
    for c,v in zip('ABCDEFGJ',row): pb[f'{c}{rr}']=v
pb['A16']='of which agreed / due this week'

# ================= CREDITORS RUN =================
for r in range(11,511):
    for c in range(1,14): cr.cell(r,c).value=None
cr['B4']=RUN; cr['C4']='Tue 1 Dec - the next monthly run (week 5). Anything with its own date in column L ignores this, and I have dated everything open at 4 Nov to its due date.'
cr['B7']=84568.84; cr['A58']=None
for i,b in enumerate(bills):
    sup,ref,due,amt,meth,cat,note,net,gst,invd,payon=b; rr=11+i
    code,_,desc=ref.partition(' - ')
    cr[f'A{rr}']=invd; cr[f'B{rr}']='Bill'; cr[f'C{rr}']=code; cr[f'E{rr}']=f'{sup} - {desc}'; cr[f'G{rr}']=net; cr[f'H{rr}']=gst; cr[f'I{rr}']=amt; cr[f'J{rr}']=amt; cr[f'K{rr}']='Awaiting Payment'
    if meth=='Card': cr[f'M{rr}']='x'
    else: cr[f'L{rr}']=payon
cr['A41']='Rows 11 to 39 are the 29 open bills in Xero at 4 Nov ($84,568.84) typed in the report layout. Next month, clear them and paste the new report over the top.'
cr['A41'].font=Font(italic=True,size=9)

# ================= CASH FLOW NOTES / RISKS =================
cf['A69']="1. Tax is the big block - Oct payroll tax $8.6k on Fri 6 Nov, the Oct IAS ~$45k on Fri 20 Nov and the Q1 BAS ~$96k on Wed 25 Nov. About $150k out in weeks 2 to 4. The BAS did not go on 28 Oct, so confirm the agent date and the GST figure off the activity statement."
cf['A70']="2. CBA Place South $118,250 (production kit, due 15 Nov) is the biggest receipt and it is not in yet. I have it Mon 16 Nov (week 3). Deloitte's Oct contract invoice ($117k) is the other big one - they have been paying in about a week, so week 2."
cf['A71']="3. Overdue debtors: Downer $39,664 (six weeks), Microsoft $13,846 (two months), Ventia $10,016, ASX $5,748, ON24 $3,006 and a few small new customers - about $85k. I have them landing in weeks 2 to 4. If none of it comes, closing is ~$85k lower. Ventia did pay $26.7k on 2 Nov, so they are moving."
cf['A72']="4. The 3 Nov pay run was $115k gross - the biggest this year - and the next two are in weeks 3 and 5. With super and payroll tax that is ~$105k each fortnight. Cochlear's $70k AGM invoice is due 4 Dec, two days after the forecast ends - upside if it comes early."

# ================= SUMMARY SHEET =================
sm=wb.create_sheet('Summary',0)
NAVY='FF002060'; GREY='FFF2F2F2'; YEL='FFFFF2CC'
hfont=Font(name='Calibri',size=11,bold=True,color='FFFFFFFF'); bfont=Font(name='Calibri',size=11); bold=Font(name='Calibri',size=11,bold=True)
money=cf['B13'].number_format
sm.column_dimensions['A'].width=46
for col in 'BCDEFG': sm.column_dimensions[col].width=15
sm['A1']='Corporate Technology Services Pty Ltd'; sm['A1'].font=Font(name='Calibri',size=14,bold=True,color=NAVY)
sm['A2']='Cash flow summary - five weeks from Thu 29 Oct to Wed 2 Dec 2026'; sm['A2'].font=Font(name='Calibri',size=12,bold=True)
sm['A3']='Prepared 4 Nov 2026 from Xero. Week 1 is actual, weeks 2 to 5 are forecast. Bank balances are cheque + savings.'; sm['A3'].font=Font(name='Calibri',size=10,italic=True)
def header(row,text,cols=7):
    for c in range(1,cols+1):
        cell=sm.cell(row,c); cell.fill=PatternFill('solid',fgColor=NAVY); cell.font=hfont
    sm.cell(row,1,text)
def line(row,label,formula,fmt=money,b=False,fill=None):
    sm.cell(row,1,label).font=bold if b else bfont
    c=sm.cell(row,2,formula); c.number_format=fmt; c.font=bold if b else bfont
    if fill:
        for cc in range(1,3): sm.cell(row,cc).fill=PatternFill('solid',fgColor=fill)
# headline
header(5,'WHERE WE ARE')
line(6,'Bank at Wed 4 Nov (cheque + savings, worked back from Xero)',"='Cash Flow'!C59",b=True,fill=YEL)
line(7,'Forecast bank at Wed 2 Dec',"='Cash Flow'!J58",b=True)
line(8,'Lowest forecast week-end balance',"=MIN('Cash Flow'!D58,'Cash Flow'!F58,'Cash Flow'!H58,'Cash Flow'!J58)")
line(9,'Receipts expected in the next four weeks',"=SUM('Cash Flow'!D15,'Cash Flow'!F15,'Cash Flow'!H15,'Cash Flow'!J15)")
line(10,'Payments expected in the next four weeks',"=SUM('Cash Flow'!D55,'Cash Flow'!F55,'Cash Flow'!H55,'Cash Flow'!J55)")
line(11,'Open customer invoices in Xero at 4 Nov',"=Receivables!F67")
line(12,'Open supplier bills in Xero at 4 Nov',"='Creditors Run'!B6")
# week table
header(14,'WEEK BY WEEK')
for i,(h) in enumerate(['','Week 1 actual','Week 2','Week 3','Week 4','Week 5','Total']):
    c=sm.cell(15,i+1,h); c.font=bold; c.fill=PatternFill('solid',fgColor=GREY); c.alignment=Alignment(horizontal='center',wrap_text=True)
for i,col in enumerate('BDFHJ'):
    c=sm.cell(16,i+2,f"='Cash Flow'!{col}6"); c.font=Font(name='Calibri',size=10,italic=True); c.alignment=Alignment(horizontal='center')
sm.cell(16,1,'Dates').font=bfont
rows=[('Opening bank',10),('Receipts',15),('Wages, super, payroll tax and the Oct IAS',24),('Job costs (equipment, hires, sub-contractors)',32),('Overheads incl. creditors run',49),('BAS and IAS',53),('Net cash movement',57),('Closing bank',58)]
rr=17
srccols=['C','D','F','H','J']
for label,cfrow in rows:
    sm.cell(rr,1,label).font=bold if cfrow in (57,58) else bfont
    for i,col in enumerate(srccols):
        c=sm.cell(rr,i+2,f"='Cash Flow'!{col}{cfrow}"); c.number_format=money; c.font=bold if cfrow in (57,58) else bfont
    if cfrow in (15,24,32,49,53,57):
        c=sm.cell(rr,7,f"=SUM(B{rr}:F{rr})"); c.number_format=money; c.font=bold if cfrow==57 else bfont
    elif cfrow==58:
        c=sm.cell(rr,7,f"=F{rr}"); c.number_format=money; c.font=bold
    elif cfrow==10:
        c=sm.cell(rr,7,f"=B{rr}"); c.number_format=money
    rr+=1
sm.cell(rr,1,'The wages line has payroll tax $8.6k in week 2 and the Oct IAS $44.7k in week 4. The BAS line is the Q1 BAS on 25 Nov.').font=Font(name='Calibri',size=9,italic=True)
# what happened
rr+=2; header(rr,'WHAT HAPPENED IN OCTOBER (28 Sep to 4 Nov)'); rr+=1
hist=[
 "Aware Super paid the $225,887 seminar room invoice on 13 Oct and everything else by 27 Oct. That was the big risk in the last forecast and it is cleared.",
 "Complete AV was paid in full - materials $139k on 29 Sep and the held labour $39k on 8 Oct - ahead of the Aware money landing.",
 "Three pay runs went out: 6 Oct $79.7k net, 20 Oct ~$88k net and 3 Nov $91k net. Gross is running at $112k to $115k a fortnight, up on the $111k assumed.",
 "The 2 Nov creditors run cleared $89.5k including the six Brookvale insurance renewals ($66k), Contact Group, Spacera and Easton Belle. Appspace ($19k) and Start Webcast ($12.8k) did not go and are in this week.",
 "The Q1 BAS did not go out on 28 Oct. It is now in the forecast on 25 Nov through the agent - needs confirming.",
 "Receipts over the five weeks to 4 Nov were about $763k against $616k forecast - Aware, and the CBA and Deloitte contract invoices landing on time.",
]
for t in hist:
    sm.cell(rr,1,t).font=bfont; sm.cell(rr,1).alignment=Alignment(wrap_text=True,vertical='top'); sm.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=7); sm.row_dimensions[rr].height=32; rr+=1
# what is coming
rr+=1; header(rr,'WHAT IS COMING IN'); rr+=1
for c,h in enumerate(['Customer','Amount','Expected','Week','Note'],1):
    cell=sm.cell(rr,c,h); cell.font=bold; cell.fill=PatternFill('solid',fgColor=GREY)
rr+=1
big_in=[(7,'CBA production kit - not in yet, due 15 Nov'),(5,'Deloitte Oct contract'),(26,'ASX AGM'),(30,'TLC AGM variation'),(29,'April 5 - new customer'),(8,'CBA Oct contract'),(21,'Aware - three Oct invoices, $31.6k together'),(33,'Downer - six weeks overdue, chasing'),(24,'Cochlear AGM - due 4 Dec, just outside the forecast')]
for recrow,note in big_in:
    sm.cell(rr,1,f'=Receivables!A{recrow}').font=bfont
    c=sm.cell(rr,2,f'=Receivables!F{recrow}'); c.number_format=money
    c=sm.cell(rr,3,f'=Receivables!H{recrow}'); c.number_format='d mmm'
    sm.cell(rr,4,f'=Receivables!J{recrow}')
    sm.cell(rr,5,note).font=bfont; sm.merge_cells(start_row=rr,start_column=5,end_row=rr,end_column=7)
    rr+=1
# what is going out
rr+=1; header(rr,'WHAT IS GOING OUT'); rr+=1
for c,h in enumerate(['Payment','Amount','When','Week','Note'],1):
    cell=sm.cell(rr,c,h); cell.font=bold; cell.fill=PatternFill('solid',fgColor=GREY)
rr+=1
outs=[('Pay run (net) - Tue 17 Nov','=Payroll!E6','=Payroll!C6','3','Plus super the day after'),
      ('Pay run (net) - Tue 1 Dec','=Payroll!E7','=Payroll!C7','5','Plus super the day after'),
      ('Q1 BAS - through the agent','=Assumptions!B22+Assumptions!B23+Assumptions!B24','=Assumptions!B25','4','GST is an estimate - swap in the activity statement figure'),
      ('October IAS (PAYG withholding)','=Assumptions!B50','=Assumptions!B51','4','21 Nov is a Saturday'),
      ('Dec rent - Investa','=Payables!D65','=Payables!G65','5','Billed 1 Dec, paid on the day last two months'),
      ('Appspace (Aware licensing)','=Payables!D19','=Payables!G19','2','Overdue - Aware already paid us'),
      ('Headroom AV - three bills','=SUM(Payables!D21:D23)','=Payables!G21','2','Check INV15 and INV-0033 are not the same job'),
      ('First Focus IT - CORE + Azure','=Payables!D52+Payables!D53','=Payables!G52','3','Direct debit'),
      ('Start Webcast - CBA AGM','=Payables!D20','=Payables!G20','2','DL approved'),
      ('October payroll tax','=Assumptions!B17+Assumptions!B18','=Assumptions!B19','2','Due 7 Nov, a Saturday')]
for label,amt,when,wk,note in outs:
    sm.cell(rr,1,label).font=bfont
    c=sm.cell(rr,2,amt); c.number_format=money
    c=sm.cell(rr,3,when); c.number_format='d mmm'
    sm.cell(rr,4,int(wk))
    sm.cell(rr,5,note).font=bfont; sm.merge_cells(start_row=rr,start_column=5,end_row=rr,end_column=7)
    rr+=1
# risks
rr+=1; header(rr,'KEY RISKS'); rr+=1
for k in (69,70,71,72):
    sm.cell(rr,1,f"='Cash Flow'!A{k}").font=bfont; sm.cell(rr,1).alignment=Alignment(wrap_text=True,vertical='top'); sm.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=7); sm.row_dimensions[rr].height=48; rr+=1
# actions
rr+=1; header(rr,'ACTIONS THIS WEEK'); rr+=1
acts=["Chase Downer ($39.7k), Microsoft ($13.8k), ASX test days ($5.7k) and the small overdue new customers - about $85k sitting past due.",
      "Confirm with the agent when the Q1 BAS is lodged and paid, and get the GST figure. The forecast has $96k on 25 Nov.",
      "Pay Appspace and Start Webcast on Fri 6 Nov, and payroll tax the same day.",
      "Check the two Headroom foldback monitor bills (INV15 and INV-0033) - same amount, same job description.",
      "Type the CommBiz cheque + savings balance at 4 Nov into the Cash Flow tab (C59) and the opening balance on Assumptions (B6). Both are worked back from Xero at the moment."]
for t in acts:
    sm.cell(rr,1,t).font=bfont; sm.cell(rr,1).alignment=Alignment(wrap_text=True,vertical='top'); sm.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=7); sm.row_dimensions[rr].height=32; rr+=1
sm.sheet_view.showGridLines=False
sm.page_setup.fitToWidth=1
wb.active=0
wb.calculation.fullCalcOnLoad=True
wb.save('out/nov4.xlsx'); print('saved')
