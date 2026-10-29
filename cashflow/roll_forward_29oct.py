import openpyxl, datetime, copy, re
from openpyxl.styles import PatternFill, Font
D=datetime.datetime
wb=openpyxl.load_workbook('out/base.xlsx')
cf=wb['Cash Flow']; a=wb['Assumptions']; p=wb['Payroll']; rec=wb['Receivables']; pb=wb['Payables']; cr=wb['Creditors Run']
NOFILL=PatternFill(fill_type=None)
def wk(d):  # week number for a date in the new table
    s=D(2026,10,29); n=(d-s).days//7+1
    return n if 1<=n<=5 else ('Before' if n<1 else 'After')

# ================= ASSUMPTIONS =================
a['A5']='Forecast start'; a['B5']=D(2026,10,29)
a['C5']="Thu 29 Oct. Five weeks of Thu to Wed, running on from the last forecast (its week 5 is this week 1). Week table is at the bottom of this tab."
a['C44']='=B44+6'
a['B6']='=896676.95-456746.27'
a['C6']="Estimate - Xero total cash at Wed 28 Oct ($896,677) less the other accounts ($456,746, the gap on 5 Oct when cheque + savings was $358,696). Replace with the CommBiz cheque + savings balance at 28 Oct."
a['B9']=107500
a['C9']='Average of the last four fortnights: the two Sep runs $217,062, 6 Oct $101,246 and 20 Oct $111,704 (Xero cash P&L). Detail on the Payroll tab.'
a['B10']=0.788
a['C10']='6 Oct run: $101,246 gross, about $79,736 left the bank = 78.8%. Was 76.8% on the 25 Aug ABA. Swap in the ABA figure when you have it.'
a['A14']='Pay date - run 1 (fortnight Sat 17 Oct to Fri 30 Oct)'; a['B14']=D(2026,11,3); a['C14']='Tue 3 Nov (week 1). Super direct debits the day after.'
a['A15']='Pay date - run 2 (fortnight Sat 31 Oct to Fri 13 Nov)'; a['B15']=D(2026,11,17); a['C15']='Tue 17 Nov (week 3).'
a['A16']='Pay date - run 3 (fortnight Sat 14 Nov to Fri 27 Nov)'; a['B16']=D(2026,12,1); a['C16']='Tue 1 Dec (week 5).'
a['A17']='Payroll tax NSW - October return (due 7 Nov)'; a['B17']=5902.75
a['C17']="Sep return was $8,602.75 all up, paid 7 Oct. I've used the same total. The NSW / other split is assumed - I can't see it in Xero."
a['A18']='Payroll tax other states - October (due 7 Nov)'; a['B18']=2700; a['C18']='Assumed, see above.'
a['B19']=D(2026,11,6); a['C19']='7 Nov is a Saturday, so Fri 6 Nov (week 2).'
a['A21']='Q1 BAS paid inside this forecast? (1 = yes, 0 = no)'; a['B21']=1
a['C21']="It did not go out on 28 Oct (Xero bank is flat 28 to 29 Oct), so I've assumed it goes through the agent and is paid 25 Nov (week 4). Change the date in B25 if it goes earlier, or set this to 0."
a['C22']='PAYG % x Sep gross $217,062 (Xero cash P&L). Jul and Aug went on the IAS, Sep goes on the Q1 BAS.'
a['A25']='BAS payment date'; a['B25']=D(2026,11,25); a['C25']='Wed 25 Nov - agent lodgement date.'
a['A39']='Invoices raised and collected inside the 5 weeks (not listed)'; a['C39']='Not used in the formulas - the Receivables tab lists everything. Left for reference.'
a['C40']='Sep was $648.92, Oct $430.17. Lands at month end - 30 Oct (week 1) and 30 Nov (week 5).'
a['A50']='October PAYG withholding - monthly IAS'; a['B50']='=B11*(101245.94+111704.13)'
a['C50']='PAYG % x the two Oct runs ($212,950 gross, Xero cash P&L). Jul and Aug were paid 21 Aug and 21 Sep on the IAS.'
a['A51']='IAS payment date'; a['B51']=D(2026,11,20); a['C51']='21 Nov is a Saturday, so Fri 20 Nov (week 4).'
for r in (50,51):
    a[f'A{r}']._style=copy.copy(a['A40']._style); a[f'B{r}']._style=copy.copy(a['B40']._style); a[f'C{r}']._style=copy.copy(a['C40']._style)
a['B51'].number_format=a['B25'].number_format; a['B50'].number_format=a['B22'].number_format

# ================= PAYROLL =================
p['A5']='Run 1 - Sat 17 Oct to Fri 30 Oct'; p['B5']=D(2026,10,30)
p['J5']='Timesheets close Fri 30 Oct, processed Mon 2 Nov, out of the bank Tue 3 Nov. Last run of the month - accrue bonus leave (Checklist step 48).'
p['A6']='Run 2 - Sat 31 Oct to Fri 13 Nov'; p['B6']=D(2026,11,13); p['J6']='Paid Tue 17 Nov.'
p['A7']='Run 3 - Sat 14 Nov to Fri 27 Nov'; p['B7']=D(2026,11,27); p['J7']='Paid Tue 1 Dec. Last run of November - accrue bonus leave again.'
p['A9']='Payroll tax for October wages (paid November)'; p['A10']='NSW - due 7 Nov (paid Fri 6 Nov)'; p['A11']='Other states - due 7 Nov (paid Fri 6 Nov)'
p['A13']='Payroll tax on Nov wages is due 7 Dec - after this forecast.'
p.insert_rows(22,2)
for r in (22,23):
    for c in range(1,7): p.cell(r,c)._style=copy.copy(p.cell(21,c)._style)
p['A22']='FE 2 Oct (paid Tue 6 Oct)'; p['B22']=D(2026,10,9); p['C22']=101245.94; p['D22']=13083.54; p['E22']='=IF(C22=0,0,D22/C22)'; p['F22']='Xero cash P&L, journal dated 8-9 Oct. Bank shows about $79.7k net leaving 6 Oct, super + payroll tax 7 Oct.'
p['A23']='FE 16 Oct (paid Tue 20 Oct)'; p['B23']=D(2026,10,20); p['C23']=111704.13; p['D23']=14214.47; p['E23']='=IF(C23=0,0,D23/C23)'; p['F23']='Xero cash P&L 13-28 Oct: all salaries & wages accounts; direct + indirect super.'
p['B24']='=(C21+C22+C23)/4'; p['C24']='The two Sep runs plus the two Oct runs, over four fortnights.'
p['D29']='=B29+C29'; p['D30']='=B30+C30'
p['A31']="Sep return paid 7 Oct was $8,602.75 all up (Xero cash P&L). Split by state not visible, so the Nov payment uses that total with the Aug split."

# ================= RECEIVABLES =================
for r in range(5,66):
    for c in range(1,12):
        cell=rec.cell(r,c); cell.value=None; cell.fill=NOFILL
base=[copy.copy(rec.cell(36,c)._style) for c in range(1,12)]
def rrow(r,vals,note):
    for c in range(1,12): rec.cell(r,c)._style=copy.copy(base[c-1])
    for c,v in zip('ABCDEFGH',vals): rec[f'{c}{r}']=v
    rec[f'I{r}']=f'=H{r}-E{r}'
    rec[f'J{r}']=f'=IF(H{r}<Assumptions!$B$5,"Before",IF(H{r}>Assumptions!$C$48,"After",INT((H{r}-Assumptions!$B$5)/7)+1))'
    rec[f'K{r}']=note
open_inv=[
 ('ASX Operations','INV-10668','External listing test days 29.9.26',D(2026,9,30),D(2026,10,30),5747.5,'Current',D(2026,10,30),'ASX pay on the due date - the 15 Oct ones landed 15 Oct.'),
 ('ASX Operations','INV-10705','ASX AGM 22.10.2026',D(2026,10,29),D(2026,11,28),70644.76,'Current',D(2026,11,30),'Due Sat 28 Nov, call it Mon 30 Nov. The big one for ASX.'),
 ('ASX Operations','INV-10704','CEO Connect series 27.10.2026',D(2026,10,29),D(2026,11,28),4041.4,'Current',D(2026,11,30),'With the AGM one.'),
 ('ASX Operations','INV-10703','Listing event lighting 14.10.26',D(2026,10,29),D(2026,11,28),1903.0,'Current',D(2026,11,30),'With the AGM one.'),
 ('April 5','INV-10706','Microsoft SME&C event 14.10.2026 (PO016066)',D(2026,10,29),D(2026,11,28),32452.77,'Current',D(2026,11,30),'New customer - on the due date, Mon 30 Nov.'),
 ('Chorus Call Australia','INV-10678','Redox AGM 07.10.2026',D(2026,10,16),D(2026,11,15),5326.2,'Current',D(2026,11,16),'Due Sun 15 Nov so Mon 16 Nov.'),
 ('Commonwealth Bank (CBA Place South)','INV-10677','Bankwest production kit (PO3311640)',D(2026,10,16),D(2026,11,15),118250.0,'Current',D(2026,11,6),'Biggest one in the forecast. CBA paid the Sep contract invoices 15 days after we sent them, so Fri 6 Nov. If they take the full 30 days it is Mon 16 Nov - still inside the forecast.'),
 ('Commonwealth Bank (CBA Place South)','INV-10707','Foldback monitor hire 22.10.2026',D(2026,10,29),D(2026,11,28),3498.0,'Current',D(2026,11,13),'15 days after invoice.'),
 ('Dimensional Fund Advisors','INV-10692','AV audio support 15.10.26',D(2026,10,20),D(2026,11,19),418.0,'Current',D(2026,11,6),'Paid the last one 9 days after invoice.'),
 ('Downer EDI','INV-10587','Full year results briefing 20.08.2026',D(2026,8,25),D(2026,9,24),12808.69,'Overdue 1 month',D(2026,11,13),'Five weeks overdue at 29 Oct and nothing in yet. Chase. Fri 13 Nov.'),
 ('Downer EDI','INV-10611','Full year employee briefing 25.08.2026',D(2026,8,31),D(2026,9,30),26855.13,'Overdue',D(2026,11,13),'A month overdue. Chase with the other one.'),
 ('Garvan Research Foundation','INV-10679','Breast cancer public lecture 01.10.2026',D(2026,10,15),D(2026,11,14),418.0,'Current',D(2026,11,16),'Due Sat 14 Nov so Mon 16 Nov.'),
 ('Garvan Research Foundation','INV-10680','Breast cancer symposium (PO-59639)',D(2026,10,15),D(2026,11,14),836.0,'Current',D(2026,11,16),'Same.'),
 ('Garvan Research Foundation','INV-10693','Suttons cocktail event 15.10.2026',D(2026,10,20),D(2026,11,19),1034.0,'Current',D(2026,11,19),'On the due date.'),
 ('Generation-e Productivity Solutions','INV-10532','PAU26-407 (Jul)',D(2026,7,31),D(2026,8,30),11850.16,'Overdue 2 months',D(2026,11,20),'Two months overdue. The Sep one was settled with a credit note on 20 Oct - check this one is not in dispute too. Doubtful.'),
 ('Global Philanthropic','INV-10670','Talking Philanthropy Aus 18.09.2026',D(2026,9,30),D(2026,10,30),4343.9,'Current',D(2026,11,6),'New customer - give it a week past due.'),
 ('Global Philanthropic','INV-10669','Talking Philanthropy event B-roll',D(2026,9,30),D(2026,10,30),1373.9,'Current',D(2026,11,6),'With the other one.'),
 ('ICC Sydney','INV-10672','Speaker presentation technicians Sep (TTS)',D(2026,9,30),D(2026,10,30),4100.26,'Current',D(2026,10,30),'ICC paid the other Sep one two weeks early. On the due date.'),
 ('Jaazaniah Salanoa (staff)','INV-10621','Goget hire recovery',D(2026,8,31),D(2026,9,30),66.64,'Overdue',D(2026,12,4),'Staff recovery - goes through payroll, no cash. Left out.'),
 ('Meet Magic','INV-10464','Human AI event highlights videos 29.05.2026',D(2026,6,19),D(2026,6,26),5075.07,'Overdue 4 months',D(2026,12,18),'DOUBTFUL - four months overdue. Left out. Needs escalating.'),
 ('Microsoft','INV-10554','PO 10160645',D(2026,7,31),D(2026,8,30),13845.7,'Overdue 2 months',D(2026,11,20),'Two months overdue - usually a PO / portal issue. Chase. Week 4.'),
 ('MUFG Corporate Markets','INV-10709','IMB Limited AGM 27.10.2026',D(2026,10,29),D(2026,11,28),9129.73,'Current',D(2026,11,30),'On the due date, Mon 30 Nov.'),
 ('Noisebox Integrated Technologies','INV-10681','Audio dry hire 13.10.26',D(2026,10,16),D(2026,11,15),203.5,'Current',D(2026,11,16),'Due Sun 15 Nov so Mon 16 Nov.'),
 ('ON24','INV-10589','PO25962',D(2026,8,26),D(2026,9,25),3005.75,'Overdue 1 month',D(2026,11,20),'The Aug one has been paid. This one is a month overdue - chase.'),
 ('PwC Services','INV-10655','PO2601702104 - Sep contract',D(2026,9,30),D(2026,10,30),13097.7,'Current',D(2026,10,30),'PwC pay on the due date (the Aug ones came 29 Sep).'),
 ('PwC Services','INV-10641','PO2601702371 - Sep contract',D(2026,9,30),D(2026,10,30),7865.0,'Current',D(2026,10,30),'On the due date.'),
 ("St Luke's Liverpool",'INV-10676','Digital mixer equipment',D(2026,9,30),D(2026,10,30),4988.5,'Current',D(2026,11,6),'New customer - give it a week past due.'),
 ('The Lottery Corporation','INV-10682','AGM event photography 14.10.2026',D(2026,10,15),D(2026,11,14),1534.5,'Current',D(2026,11,16),'Due Sat 14 Nov so Mon 16 Nov.'),
 ('The Lottery Corporation','INV-10710','AGM 14.10.2026 (variation)',D(2026,10,29),D(2026,11,28),39187.49,'Current',D(2026,11,30),'New customer - on the due date, Mon 30 Nov.'),
 ('Vega Global Australia','INV-10647','Payments & Securities Forums 14-17.9.26',D(2026,9,22),D(2026,10,22),1969.0,'Overdue',D(2026,11,6),'A week overdue - chase.'),
 ('Vega Global Australia','INV-10684','AI Forum audio assistance 13-14.10.26',D(2026,10,16),D(2026,11,15),1907.13,'Current',D(2026,11,16),'Due Sun 15 Nov so Mon 16 Nov.'),
 ('Ventia','INV-10612','Half year results briefing 24.08.2026',D(2026,8,31),D(2026,9,30),11919.04,'Overdue',D(2026,11,13),'A month overdue. Ventia paid the big Aug one on 29 Sep, so these should come - chase. Fri 13 Nov.'),
 ('Ventia','INV-10593','Employee HYR town hall 24.08.2026',D(2026,8,31),D(2026,9,30),14741.65,'Overdue',D(2026,11,13),'Same.'),
 ('Ventia','INV-10632','RUOK Day event 10.09.2026',D(2026,9,16),D(2026,10,16),10015.89,'Overdue',D(2026,11,13),'Same.'),
 ('YSG Studio','INV-10653','YSG Studio phase 4',D(2026,9,25),D(2026,10,25),1116.5,'Overdue',D(2026,11,6),'A few days overdue - chase.'),
]
tbr=[
 ('Deloitte Services Trust','TBR','Oct contract invoice (template $113,988)',D(2026,10,30),D(2026,11,29),113987.99,'To be raised',D(2026,11,9),'Sep contract invoice (30 Sep) was paid 8 Oct - Deloitte pay about a week after we invoice.'),
 ('Deloitte Services Trust','TBR','Oct additionals (Sep actual $18,892 used)',D(2026,10,30),D(2026,11,29),18891.66,'To be raised',D(2026,11,9),'Sep additionals were paid with the contract one. Amount will differ.'),
 ('Commonwealth Bank (CBA Place South)','TBR','Oct contract (85720001CBPSCSJUN27)',D(2026,10,30),D(2026,11,29),24619.84,'To be raised',D(2026,11,16),'Sep ones (30 Sep) were paid 15 Oct - 15 days.'),
 ('Commonwealth Bank (CBA Place South)','TBR','Oct dedicated tech / after hours (Sep actual $3,432)',D(2026,10,30),D(2026,11,29),3432.0,'To be raised',D(2026,11,16),'Same.'),
 ('Commonwealth Bank (CBA Place South)','TBR','Oct live event support (Sep actual $5,398)',D(2026,10,30),D(2026,11,29),5398.25,'To be raised',D(2026,11,16),'Same.'),
 ('Commonwealth Bank (CBA Square)','TBR','Oct contract (85720001CBSCSJUN27)',D(2026,10,30),D(2026,11,29),11792.0,'To be raised',D(2026,11,16),'Same.'),
 ('Commonwealth Bank (CBA Melbourne)','TBR','Oct contract (85720001MEL435CSJUN27)',D(2026,10,30),D(2026,11,29),12995.72,'To be raised',D(2026,11,16),'Same.'),
 ('Commonwealth Bank (CBA Brisbane)','TBR','Oct contract (85720001BRISCSJUN27, Sep actual)',D(2026,10,30),D(2026,11,29),8043.2,'To be raised',D(2026,11,9),'Sep one was paid 7 Oct - a week.'),
 ('Bankwest','TBR','Oct contract (85720001BWPCSJUN27)',D(2026,10,30),D(2026,11,29),13156.0,'To be raised',D(2026,11,3),'Bankwest paid the Sep one 2 Oct - two days after we sent it.'),
 ('Bankwest','TBR','Oct additionals (Sep actual $4,752)',D(2026,10,30),D(2026,11,29),4752.0,'To be raised',D(2026,11,16),'Sep additionals took two weeks.'),
 ('PwC Services','TBR','Oct contract PO2601702104',D(2026,10,30),D(2026,11,29),13097.7,'To be raised',D(2026,11,30),'PwC pay on the due date. Due Sun 29 Nov so Mon 30 Nov.'),
 ('PwC Services','TBR','Oct contract PO2601702371',D(2026,10,30),D(2026,11,29),7865.0,'To be raised',D(2026,11,30),'Same.'),
 ('Aware Super','TBR','Annual National Maintenance - October',D(2026,10,30),D(2026,11,29),12100.0,'To be raised',D(2026,11,27),'Aware paid the Sep one 27 Oct, three days before due.'),
]
r=5
for row in open_inv+tbr:
    rrow(r,row[:8],row[8]); r+=1
print('receivable rows used to',r-1)
rec['A2']=("Xero aged receivables at 29 Oct ($445,570.46, 35 invoices) plus the Oct contract invoices still to be raised (from the repeating templates, Sep actuals used where the amount varies). "
           "Expected date (blue) is my call based on how each customer has been paying - change it and the forecast moves. Aware Super paid everything in October, including the $225,887.")
rec['A67']='of which open in Xero at 29 Oct (should be $445,570.46)'
rec['F67']='=SUMIFS(F5:F65,G5:G65,"<>To be raised")'
rec['A68']='of which expected inside the forecast (29 Oct to 2 Dec)'
rec['A69']='of which pushed past the forecast (Meet Magic, Jaaz, the Nov contract invoices are not listed)'

# ================= PAYABLES =================
# extend the list area to row 95: move totals 77-80 down to 97-100
for r in range(77,81):
    for c in range(1,16):
        s=pb.cell(r,c); d=pb.cell(r+20,c); d.value=s.value; d._style=copy.copy(s._style); s.value=None; s.fill=NOFILL
def fix75(f):
    if not isinstance(f,str) or not f.startswith('='): return f
    return re.sub(r'(\$?[DEFGIK]\$?19:\$?[DEFGIK]\$?)75\b', r'\g<1>95', f)
for ws in (cf,pb):
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value,str) and c.value.startswith('='): c.value=fix75(c.value)
tmpl=[copy.copy(pb.cell(60,c)._style) for c in range(1,16)]
for r in range(19,96):
    for c in range(1,16):
        cell=pb.cell(r,c); cell.value=None; cell._style=copy.copy(tmpl[c-1])
    pb[f'H{r}']=f'=IF(G{r}<Assumptions!$B$5,"Before",IF(G{r}>Assumptions!$C$48,"After",INT((G{r}-Assumptions!$B$5)/7)+1))'
    pb[f'K{r}']=f'=IF(AND(\'Creditors Run\'!$B$5=1,I{r}="Xero bill",LEFT(J{r},4)<>"PAID"),"No - in the creditors run","Yes")'
RUN=D(2026,11,2)
# (supplier, bill ref, due, amount, method, category, notes, net, gst) - Xero bills open at 29 Oct
bills=[
 ('Brookvale Insurance Brokers','INV1918462 - IT liability insurance (annual)',D(2026,10,30),35304.46,'EFT','Insurance','GC approved. In the 2 Nov run.',32254.45,3050.02),
 ('Brookvale Insurance Brokers','INV1918336 - removable / electronic equipment insurance (annual)',D(2026,10,30),11283.16,'EFT','Insurance','GC approved. In the 2 Nov run.',10300.62,982.54),
 ('Brookvale Insurance Brokers','INV1918436 - Mecon contract works insurance (annual)',D(2026,10,30),8020.39,'EFT','Insurance','GC approved. In the 2 Nov run.',7338.73,681.66),
 ('Brookvale Insurance Brokers','INV1920058 - management liability insurance (annual)',D(2026,10,30),6778.21,'EFT','Insurance','GC approved. In the 2 Nov run.',6195.64,582.57),
 ('Brookvale Insurance Brokers','INV1920113 - business pack insurance (annual)',D(2026,10,30),3510.52,'EFT','Insurance','GC approved. In the 2 Nov run.',3211.96,298.57),
 ('Brookvale Insurance Brokers','INV1914759 - corporate travel insurance (annual)',D(2026,10,30),1342.24,'EFT','Insurance','GC approved. In the 2 Nov run.',1246.49,95.75),
 ('Appspace','INV00133660 - Appspace licensing for Aware Super (balance)',D(2026,10,24),19077.66,'EFT','Client subscriptions','Overdue - Aware paid us for it 20 Oct. $19,077.66 left on it in Xero. In the 2 Nov run.',17343.33,1734.33),
 ('Start Webcast','INV-202222 - CBA AGM 2026 equipment & services',D(2026,10,29),12787.06,'EFT','Equipment & service hires','DL / TW still to approve. In the 2 Nov run once approved.',11624.6,1162.46),
 ('Contact Group','INV457215 - DTTL Hobart AV fitout equipment & labour',D(2026,11,1),9652.54,'EFT','Equipment purchases','In the 2 Nov run.',8775.0,877.54),
 ('Contact Group','INV457214 - DTTL Hobart AV fitout equipment',D(2026,11,1),903.38,'EFT','Equipment purchases','In the 2 Nov run.',821.25,82.13),
 ('SPACERA','INV-0060 - Aware Super national support October',D(2026,10,31),5728.8,'EFT','Client subscriptions','In the 2 Nov run.',5208.0,520.8),
 ('Easton Belle','INV00004009 - legal fees, general commercial file',D(2026,10,30),3932.5,'EFT','Professional fees','DL approved. In the 2 Nov run.',3575.0,357.5),
 ('Headroom Audio Visual','INV15 - CBA foldback monitor dry hire 21.10.26',D(2026,10,30),2992.0,'EFT','Equipment & service hires','MH approved. In the 2 Nov run.',2720.0,272.0),
 ('Hire Intelligence','CS167608 - SUN investor day 29.10.26 equipment hire',D(2026,10,30),2585.0,'EFT','Equipment & service hires','MH approved. In the 2 Nov run.',2350.0,235.0),
 ('2020 Events','INV-1611 - Microsoft SME&C townhall 14.10.26 hire',D(2026,11,4),2346.37,'EFT','Equipment & service hires','MH approved, needs the cost split. In the 2 Nov run.',2133.06,213.31),
 ('Employsure','42 of 60 - October',D(2026,10,5),880.0,'EFT','Professional fees','Overdue. In the 2 Nov run.',800.0,80.0),
 ('Type Transcripts','INV6062 - TLC AGM transcription 14.10.2026',D(2026,10,30),257.4,'EFT','Equipment & service hires','MH approved. In the 2 Nov run.',234.0,23.4),
 ('Ricoh Australia','INV15775235 - printing Sep',D(2026,10,31),210.2,'DD','Office','Direct debit.',191.09,19.11),
 ('2020 Events','INV-1618 - Cochlear AGM 26.10.26 hire',D(2026,11,9),5665.0,'EFT','Equipment & service hires','MH approved, needs the quote to split costs. On the due date.',5150.0,515.0),
 ('WordBirds','INV-2016 - ASX AGM 22.10.26 services',D(2026,11,9),5063.5,'EFT','Equipment & service hires','MH approved. On the due date.',4603.18,460.32),
 ('WordBirds','INV-2019 - Cochlear AGM 26.10.26 services',D(2026,11,10),3267.0,'EFT','Equipment & service hires','MH approved. On the due date.',2970.0,297.0),
 ('Dubology Studios','INV26-12398 - Microsoft SME&C and CBA LES Oct services',D(2026,11,9),2117.5,'EFT','Equipment & service hires','MH approved. On the due date.',1925.0,192.5),
 ('Deaf Services','INV409016 - CBA AGM 2026 services',D(2026,11,9),310.75,'EFT','Equipment & service hires','DL / TW to approve and confirm the cost centre.',282.5,28.25),
 ('Optus','INV000600307641 - staff mobiles Oct',D(2026,11,9),840.98,'BPAY','Telephone & internet','BPAY on the due date.',764.53,76.41),
 ('Posh Services (Ambius)','INV21742621 - indoor plants Oct',D(2026,11,8),245.0,'EFT','Office','Due Sun 8 Nov so Mon 9 Nov.',222.73,22.27),
 ('Telstra','TTI20233644300 - Nighthawk Oct',D(2026,11,6),136.0,'BPAY','Telephone & internet','BPAY on the due date.',123.64,12.36),
 ('Officeworks','INV632570641 - keyboard & mouse for CBA AGM',D(2026,11,6),108.0,'EFT','Office','DL approved. On the due date.',98.18,9.82),
 ('Riedel Communications','INV00010410 - Microsoft SME&C townhall 14.10.26 hire',D(2026,11,14),885.06,'EFT','Equipment & service hires','MH approved. Due Sat 14 Nov so Mon 16 Nov.',804.6,80.46),
 ('Access Innovation Media','INV22273 - CBA AGM 2026 services',D(2026,11,18),924.0,'EFT','Equipment & service hires','DL / TW to approve.',840.0,84.0),
 ('Access Innovation Media','INV22274 - CBA AGM 2026 services',D(2026,11,18),682.0,'EFT','Equipment & service hires','DL / TW to approve.',620.0,62.0),
 ('Access Innovation Media','INV22272 - CBA AGM 2026 services',D(2026,11,18),264.0,'EFT','Equipment & service hires','DL / TW to approve.',240.0,24.0),
 ('Access Innovation Media','INV22275 - CBA AGM 2026 services',D(2026,11,18),165.0,'EFT','Equipment & service hires','DL / TW to approve.',150.0,15.0),
 ('Riedel Communications','INV00010482 - MUFG Suncorp investor day 29.10.2026 hire',D(2026,11,25),1829.52,'EFT','Equipment & service hires','MH approved. On the due date.',1663.2,166.32),
 ('Riedel Communications','INV00010469 - ASX AGM 22.10.26 hire',D(2026,11,26),1834.27,'EFT','Equipment & service hires','MH approved. On the due date.',1667.52,166.75),
 ('Chorus Call Australia','INV02-82515 - Cochlear AGM 26.10.26 services',D(2026,11,27),2065.84,'EFT','Equipment & service hires','MH approved. On the due date.',1878.04,187.8),
 ('Ricoh Australia','INV15788514 - printing Oct',D(2026,11,30),257.47,'DD','Office','Direct debit.',234.06,23.41),
 ('Avetta','INV02434590 - worker management annual licence 2026-27',D(2026,10,23),892.0,'Card','Dues & subscriptions','PL card - in the card line.',892.0,0.0),
 ('Australian Tenders','WLVRMZBY0012 - WA core subscription Oct',D(2026,10,27),99.0,'Card','Dues & subscriptions','PL card.',90.0,9.0),
 ('Energy Australia','INV260412412804 - electricity 23 Aug to 21 Sep',D(2026,10,14),696.25,'Card','Electricity','DD via the PL card.',632.95,63.3),
 ('Expedia','IT73569434492806 - accommodation J Salanoa, Ventia townhall',D(2026,10,20),233.83,'Card','Direct travel & freight','PL card.',212.57,21.26),
 ('Virgin Australia','JGTOZQ - return flight J Salanoa, Ventia townhall',D(2026,10,20),716.0,'Card','Direct travel & freight','PL card.',650.91,65.09),
 ('Google','INV1115.7 - YouTube Premium Oct',D(2026,10,28),16.99,'Card','Dues & subscriptions','DL card.',15.45,1.54),
 ('Mailchimp','MC27578191 - October',D(2026,10,18),38.54,'Card','Dues & subscriptions','PL card.',38.54,0.0),
 ('Optus','INV000599871925 - office internet & landline Oct',D(2026,10,23),639.15,'Card','Telephone & internet','PL card.',581.04,58.11),
 ('Seek','INV702171405 - accounts assistant job ad',D(2026,10,21),555.5,'Card','Recruitment','GC approved. PL card.',505.0,50.5),
 ('Zoom','INV375651107 - Oct subscription',D(2026,10,28),323.97,'Card','Dues & subscriptions','PL card.',323.97,0.0),
]
assert len(bills)==46 and abs(sum(b[3] for b in bills)-158464.01)<0.02, sum(b[3] for b in bills)
templates=[
 ('Investa Asset Management','Rent November 2026 (Oct was $28,895.72)',D(2026,11,1),28895.72,'EFT','Rent','Repeating template','Investa bill the 1st, paid 1 Oct last month. Pay Mon 2 Nov with the run.'),
 ('Cleveland Partners (8020 Advisors)','Monthly advisory fees - November',D(2026,11,2),3850.0,'EFT','Professional fees','Repeating template','Template $3,850 a month. Nothing paid in Oct that I can see in Xero - check.'),
 ('Employsure','43 of 60 - November',D(2026,11,5),880.0,'EFT','Professional fees','Repeating template','Template, due 5 Nov.'),
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
 ('SPACERA','Aware Super national support - November',D(2026,12,1),5728.8,'EFT','Client subscriptions','Repeating template','Due Tue 1 Dec - week 5.'),
 ('Investa Asset Management','Rent December 2026',D(2026,12,1),28895.72,'EFT','Rent','Repeating template','Billed 1 Dec. Last two months were paid on the day, so it lands in week 5 (Tue 1 Dec). Push the date out if the Dec run is later.'),
 ('Posh Services (Ambius)','Indoor plants - November',D(2026,12,9),245.0,'EFT','Office','Repeating template','After this forecast.'),
 ('Optus','Staff mobiles - November',D(2026,12,9),804.06,'BPAY','Telephone & internet','Repeating template','After this forecast.'),
 ('The CEO Institute','Executive membership (JS) - quarterly',D(2026,12,7),2310.0,'EFT','Dues & subscriptions','Repeating template','After this forecast.'),
 ('AVIXA','DL annual silver membership',D(2026,12,11),2850.0,'EFT','Dues & subscriptions','Repeating template','After this forecast.'),
]
r=19
for b in bills:
    sup,ref,due,amt,meth,cat,note,net,gst=b
    payon = RUN if (due<=RUN or meth=='Card' and due<=RUN) else (due if due.weekday()<5 else due+datetime.timedelta(days=7-due.weekday()))
    if meth=='Card': payon=due
    for c,v in zip('ABCDEFGIJ',(sup,ref,due,amt,meth,cat,payon,'Xero bill',note)): pb[f'{c}{r}']=v
    r+=1
for t in templates:
    sup,ref,due,amt,meth,cat,src,note=t
    payon=due if due.weekday()<5 else due+datetime.timedelta(days=7-due.weekday())
    for c,v in zip('ABCDEFGIJ',(sup,ref,due,amt,meth,cat,payon,src,note)): pb[f'{c}{r}']=v
    r+=1
print('payables rows used to',r-1)
pb['A97']='TOTAL bills listed from Xero (aged payables at 29 Oct, should be $158,464.01)'
pb['A2']=("Xero aged payables at 29 Oct ($158,464.01, 46 bills) plus the repeating bill templates for November. Card rows get paid on the PL/DL cards and sit in the credit card line on the Cash Flow tab, so they're left out of the category totals. "
          "With the Creditors Run tab switched on, the open Xero bills come through the Creditors run row and the templates come through the category lines.")
# payments this week block (29 Oct - 4 Nov)
pb['A4']='PAYMENTS THIS WEEK (29 Oct - 4 Nov) - creditors run Mon 2 Nov'
block=[
 ('Brookvale Insurance Brokers','Six annual renewals - INV1918462, 1918336, 1918436, 1920058, 1920113, 1914759',D(2026,10,30),'=SUM(D19:D24)','EFT','AGREED - PAY',RUN,'GC approved all six. One payment in the Mon 2 Nov run - $66,238.98.'),
 ('=A25','=B25','=C25','=D25','=E25','Due this week','=G25','Appspace - Aware paid us for it 20 Oct. Overdue, pay in the run.'),
 ('=A26','=B26','=C26','=D26','=E26','AWAITING DL / TW','=G26','Start Webcast - DL / TW still to approve. Hold until they do.'),
 ('Contact Group','INV457215 + INV457214 - DTTL Hobart fitout',D(2026,11,1),'=D27+D28','EFT','Due this week',RUN,'Due Sun 1 Nov - in the run.'),
 ('=A29','=B29','=C29','=D29','=E29','Due this week','=G29','Spacera Oct - in the run.'),
 ('=A30','=B30','=C30','=D30','=E30','Due this week','=G30','Easton Belle - DL approved, in the run.'),
 ('Headroom / Hire Intelligence / 2020 Events / Employsure / Type Transcripts / Ricoh','Small ones due by 4 Nov',D(2026,11,4),'=SUM(D31:D36)','EFT','Due this week',RUN,'All MH approved except Employsure (overdue) and Ricoh (DD).'),
 ('=A65','=B65','=C65','=D65','=E65','Due this week','=G65','Investa Nov rent - template, pay with the run like last month.'),
 ('=A66','=B66','=C66','=D66','=E66','Due this week','=G66','8020 Advisors - template, check the Nov invoice has come in.'),
]
green=[copy.copy(pb.cell(6,c)._style) for c in range(1,16)]
plain=[copy.copy(pb.cell(7,c)._style) for c in range(1,16)]
amber=copy.copy(pb.cell(9,6)._style)
for i,row in enumerate(block):
    rr=6+i
    for c in range(1,16):
        pb.cell(rr,c)._style=copy.copy(green[c-1] if i==0 else plain[c-1])
        if i: pb.cell(rr,c).fill=NOFILL
    for c,v in zip('ABCDEFGJ',row): pb[f'{c}{rr}']=v
    if row[5].startswith('AWAITING'):
        for c in range(1,11): pb.cell(rr,c).fill=PatternFill('solid',fgColor='FFFFEB9C')
pb['A16']='of which agreed / due this week (excludes Start Webcast until DL / TW approve)'
pb['D16']='=SUMIFS(D6:D14,F6:F14,"AGREED - PAY")+SUMIFS(D6:D14,F6:F14,"Due this week")'

# ================= CREDITORS RUN =================
for r in range(11,511):
    for c in range(1,14): cr.cell(r,c).value=None
cr['B4']=RUN; cr['C4']='Mon 2 Nov (1 Nov is a Sunday). Change it each month. Anything with its own date in column L ignores this.'
cr['B7']=158464.01
cr['A37']=None
r=11
for b in bills:
    sup,ref,due,amt,meth,cat,note,net,gst=b
    code,_,desc=ref.partition(' - ')
    cr[f'A{r}']=due-datetime.timedelta(days=30) if False else None
    cr[f'B{r}']='Bill'; cr[f'C{r}']=code; cr[f'E{r}']=f'{sup} - {desc}'; cr[f'G{r}']=net; cr[f'H{r}']=gst; cr[f'I{r}']=amt; cr[f'J{r}']=amt; cr[f'K{r}']='Awaiting Payment'
    r+=1
# invoice dates from Xero for the paste rows
inv_dates={'INV1918462':D(2026,9,30),'INV1918336':D(2026,9,30),'INV1918436':D(2026,9,30),'INV1920058':D(2026,9,30),'INV1920113':D(2026,9,30),'INV1914759':D(2026,9,30),
 'INV00133660':D(2026,9,24),'INV-202222':D(2026,10,14),'INV457215':D(2026,9,30),'INV457214':D(2026,9,30),'INV-0060':D(2026,10,1),'INV00004009':D(2026,9,30),'INV15':D(2026,10,25),
 'CS167608':D(2026,10,29),'INV-1611':D(2026,10,21),'42 of 60':D(2026,10,5),'INV6062':D(2026,10,15),'INV15775235':D(2026,9,25),'INV-1618':D(2026,10,26),'INV-2016':D(2026,10,26),
 'INV-2019':D(2026,10,27),'INV26-12398':D(2026,10,26),'INV409016':D(2026,10,26),'INV000600307641':D(2026,10,10),'INV21742621':D(2026,10,9),'TTI20233644300':D(2026,10,25),
 'INV632570641':D(2026,10,7),'INV00010410':D(2026,10,15),'INV22273':D(2026,10,19),'INV22274':D(2026,10,19),'INV22272':D(2026,10,19),'INV22275':D(2026,10,19),'INV00010482':D(2026,10,26),
 'INV00010469':D(2026,10,27),'INV02-82515':D(2026,10,28),'INV15788514':D(2026,10,26),'INV02434590':D(2026,10,23),'WLVRMZBY0012':D(2026,10,27),'INV260412412804':D(2026,9,24),
 'IT73569434492806':D(2026,10,20),'JGTOZQ':D(2026,10,20),'INV1115.7':D(2026,10,28),'MC27578191':D(2026,10,18),'INV000599871925':D(2026,10,7),'INV702171405':D(2026,10,21),'INV375651107':D(2026,10,28)}
for i,b in enumerate(bills):
    rr=11+i; code=b[1].partition(' - ')[0]
    cr[f'A{rr}']=inv_dates[code]
    payon=pb[f'G{19+i}'].value
    if b[4]=='Card': cr[f'M{rr}']='x'; cr[f'L{rr}']=None
    elif payon!=RUN: cr[f'L{rr}']=payon
cr['A58']='Rows 11 to 56 are the 46 open bills in Xero at 29 Oct ($158,464.01) typed in the report layout. Next month, clear them and paste the new report over the top.'
cr['A58'].font=Font(italic=True,size=9)

# ================= CASH FLOW =================
cf['A2']='Weekly Cash Flow Forecast vs Actual - 29 October to 2 December 2026'
for col in 'CEGIK':
    for r in list(range(13,15))+list(range(19,24))+list(range(27,32))+list(range(35,49))+[52,59]:
        cf[f'{col}{r}'].value=None
    cf[f'{col}8']=0
cf['B14']='=Assumptions!$B$40'; cf['J14']='=Assumptions!$B$40'
for col,w in zip('BDFHJ',range(44,49)):
    cf[f'{col}21']=f'=IF(AND(Assumptions!$B$51>=Assumptions!$B${w},Assumptions!$B$51<=Assumptions!$C${w}),Assumptions!$B$50,0)'
cf['A48']='Creditors run (Xero aged payables paste - run Mon 2 Nov)'
notes={
 'O8':'Set to 1 once the week is done and the CommBiz balance is in. Then the next week opens from the real number.',
 'O13':'Receivables tab. Week 1 is PwC, ICC and ASX on the due date (30 Oct) plus the Bankwest Oct contract. Week 2 is the CBA production kit ($118k) and the Deloitte Oct invoices. Week 5 is the ASX AGM, TLC and April 5.',
 'O14':'Savings interest at month end - 30 Oct and 30 Nov.',
 'O19':'Payroll tab - gross x net %. Fortnight is Sat to Fri and the money goes out the Tuesday after: Tue 3 Nov (week 1), Tue 17 Nov (week 3) and Tue 1 Dec (week 5).',
 'O20':'Gross x 12%. Direct debits the day after pay day.',
 'O21':'Oct PAYG on the monthly IAS, due 21 Nov - paid Fri 20 Nov (week 4). Assumptions B50.',
 'O22':'Oct wages, due 7 Nov - a Saturday, so Fri 6 Nov (week 2). Sep was $8,602.75 all up, same total used.',
 'O23':'Placeholder - goes out with the ABA.',
 'O27':'Contact Group Hobart fitout is in the creditors run. This line is just the unbilled estimate now.',
 'O28':'Estimate (Sep was ~$39k, Oct ~$18k so far), spread evenly.',
 'O29':'Estimate, spread evenly. The hire bills already in Xero (WordBirds, Riedel, 2020 Events, Start Webcast) are in the creditors run.',
 'O30':'LiveU and Spacera Nov templates plus the unbilled estimate. Appspace and Spacera Oct are in the creditors run.',
 'O31':'Estimate plus the Goget template (card).',
 'O35':'Investa Nov rent ~$28,896 billed 1 Nov, paid with the 2 Nov run (week 1). Dec rent lands Tue 1 Dec (week 5) if they bill on the day again.',
 'O36':'First Focus CORE + Azure, DD 16 Nov (week 3).',
 'O37':'RJW digital marketing Nov $4,675, due 21 Nov.',
 'O38':'8020 Advisors $3,850 and Employsure $880 templates, plus the estimate. Easton Belle legal is in the creditors run.',
 'O39':'Card subscriptions sit in the card line. CEO Institute and AVIXA fall after 2 Dec.',
 'O40':'Optus Lugstein BPAY, Telstra Nighthawk Nov. Optus mobiles Oct is in the creditors run, Nov falls after 2 Dec.',
 'O41':'Estimate. Seek ads go on the card.',
 'O42':'Brookvale $66,239 is in the creditors run (2 Nov). Nothing else expected - estimate is nil.',
 'O43':'Ricoh DD and Posh are in the creditors run, plus the estimate.',
 'O44':'Energy Australia DDs the PL card so it shows in the credit card line, not here.',
 'O45':'Estimate - bank charges and Stripe.',
 'O46':'~$25k a month transferred to the PL and DL cards. Covers every Card row on the Payables tab.',
 'O47':'Nothing listed.',
 'O48':'Creditors Run tab - the 46 open bills at 29 Oct ($158k). Everything due by 2 Nov goes in the Mon 2 Nov run ($131k incl. Brookvale $66k and Appspace $19k). The rest lands on its due date. Start Webcast $12.8k is in the run but DL / TW still need to approve it.',
 'O52':'="Q1 BAS est. "&TEXT(Assumptions!$B$22+Assumptions!$B$23+Assumptions!$B$24,"$#,##0")&" (Sep PAYG + GST estimate). Did not go 28 Oct, so assumed 25 Nov through the agent (week 4). Swap the GST estimate for the activity statement figure."',
 'O58':"Opening + net. The Actual column fills in once the week's flag is 1, and the total column shows the latest completed week.",
 'O59':'Type the real closing balance (cheque + savings) in at the end of each week. The opening balance on the Assumptions tab is an estimate off Xero - replace it with the 28 Oct CommBiz figure first.',
 'O60':"Should be nil once everything's in. This is a sense check.",
 'O65':'Nov contract invoices (raised 30 Nov, paid Dec), Meet Magic, Jaaz.',
 'O66':'Posh Nov, Optus mobiles Nov, CEO Institute, AVIXA.',
}
for k,v in notes.items(): cf[k]=v
cf['A65']='   Receipts expected after 2 Dec (not in this forecast)'
cf['A66']='   Payments held past 2 Dec (not in this forecast)'
cf['A69']="1. Tax is the big block this month - Oct payroll tax $8.6k on 6 Nov, the Oct IAS ~$45k on 20 Nov and the Q1 BAS ~$96k on 25 Nov. That is ~$150k out in weeks 2 to 4. The BAS did not go on 28 Oct, so confirm when it is actually due and the GST figure off the activity statement."
cf['A70']="2. CBA Place South $118,250 (INV-10677, production kit, due 15 Nov) is the biggest receipt. CBA have been paying in about 15 days, so I have it Fri 6 Nov. If it slips to the due date it moves to week 3 - still in the forecast, but week 2 would dip."
cf['A71']="3. Overdue debtors: Downer $39,664, Ventia $36,677, Microsoft $13,846, Gen-e $11,850 and ON24 $3,006 - about $105k, all a month or more old. I have them landing in weeks 3 and 4. If none of it comes, closing is ~$105k lower. Chase this week."
cf['A72']="4. Creditors run Mon 2 Nov is ~$131k including Brookvale $66k and Appspace $19k, plus Nov rent $29k and the 3 Nov pay run ~$98k in the same week. Week 1 is the tightest week - make sure the CBA / Deloitte money is in before the run if the balance looks thin."
wb.calculation.fullCalcOnLoad=True
wb.save('out/roll.xlsx'); print('saved')
