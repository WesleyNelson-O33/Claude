import openpyxl, datetime, copy, re, sys
SRC='/tmp/claude-0/-home-user-Claude/48a2c21e-7430-596d-b786-a45eabe08678/scratchpad/in/cf.xlsx'
OUT=sys.argv[1]
D=datetime.datetime
wb=openpyxl.load_workbook(SRC)
a=wb['Assumptions']; p=wb['Payroll']; cf=wb['Cash Flow']; pb=wb['Payables']

# ---------------- 1. Payroll timing: Sat-Fri fortnight, paid the Tuesday after ----------------
a['A8']='PAYROLL (fortnightly, Sat to Fri, paid the Tuesday after the fortnight ends)'
a['C5']="Mon 28 Sep. Week 1 runs through to Wed 7 Oct so it picks up payroll tax on the 7th, then it's weekly from there (Thu to Wed). Week table is at the bottom of this tab."
a['A14']='Pay date - run 1 (fortnight Sat 26 Sep to Fri 9 Oct)'
a['B14']=D(2026,10,13)
a['C14']='Fortnight is Sat 26 Sep to Fri 9 Oct. Processed Mon 12 Oct, out of the bank Tue 13 Oct (week 2). Xero has the pay run journal dated 8-9 Oct for $101,246 gross - that is the journal date, the money goes on the Tuesday. Super direct debits the same week - payday super since 1 Jul 26.'
a['A15']='Pay date - run 2 (fortnight Sat 10 Oct to Fri 23 Oct)'
a['B15']=D(2026,10,27)
a['C15']='Tue 27 Oct (week 4).'
a['A16']='Pay date - run 3 (fortnight Sat 24 Oct to Fri 6 Nov)'
a['B16']=D(2026,11,10)
a['C16']='Tue 10 Nov - after week 5 (ends Wed 4 Nov), so it does not hit this forecast. Only two pay runs land in October now.'

p['A2']='Fortnightly, Saturday to Friday. The money only goes out of the bank on the Tuesday after the fortnight ends. Super direct debits the same week. PAYG goes later with the IAS/BAS.'
p['B4']='Fortnight ending (Fri)'
p['A5']='Run 1 - Sat 26 Sep to Fri 9 Oct'
p['B5']=D(2026,10,9)
p['J5']='Timesheets close Fri 9 Oct, processed Mon 12 Oct - approvals, reconcile, finalise with GC, ABA into CommBiz, payroll email to DL. Move the funds from savings Monday. Out of the bank Tue 13 Oct. Xero journal is dated 8-9 Oct ($101,246 gross, $13,084 super).'
p['A6']='Run 2 - Sat 10 Oct to Fri 23 Oct'
p['B6']=D(2026,10,23)
p['J6']='Paid Tue 27 Oct. Last run of the month - accrue bonus leave (Checklist step 48). No cash effect.'
p['A7']='Run 3 - Sat 24 Oct to Fri 6 Nov'
p['B7']=D(2026,11,6)
p['J7']='Paid Tue 10 Nov - after this forecast, so nothing from it hits the numbers.'
p['B16']='Journal date'

cf['O19']='Payroll tab - gross x net %. Fortnight is Sat to Fri and the money goes out the Tuesday after: Tue 13 Oct (week 2) and Tue 27 Oct (week 4). The next one is Tue 10 Nov, after this forecast.'

# ---------------- 2. Week 1 flag, week 2 actuals to Mon 12 Oct ----------------
cf['C8']=1          # week 1 ties to the bank (difference nil), so week 2 opens from the real balance
cf['O8']='Week 1 is done and ties to CommBiz. Week 2 actuals are in to Mon 12 Oct only - flag stays 0 until the week closes Wed 14 Oct and the bank balance is keyed.'
cf['E13']='=112683.22+18891.66+9086.39+1661+506+374'
cf['O13']='Week 2 to 12 Oct: Deloitte $112,683.22 and $18,891.66, CEO Institute $9,086.39, ASX $1,661, Dimensional $506, Ricoh $374 (Xero). Nothing from Aware yet.'
cf['E14']=0
# payroll - Tue 13 Oct still to come
for r in (21,22): cf[f'E{r}']=0
cf['O23']='Placeholder. Week 2 actual gets filled on pay day Tue 13 Oct.'
# cost of sales - Xero cash P&L 8-12 Oct, ex GST x 1.1
cf['E27']='=ROUND(28736.64*1.1,2)'
cf['O27']=cf['O27'].value.strip()+' Week 2 to 12 Oct: LiveU SDI/modems $26,004 plus small items ($28,737 ex GST per Xero cash P&L).'
cf['E28']='=ROUND((8125.22+520.4)*1.1,2)'
cf['O28']='Estimate (Sep was ~$39k), spread evenly. Week 2 actual to 12 Oct is Integration $8,125 + PRD $520 ex GST (Xero cash P&L).'
cf['E29']='=ROUND((11867.45+2204.89+222.68)*1.1,2)'
cf['O29']='Estimate (Sep was ~$29k), spread evenly. Week 2 to 12 Oct: Amora Hotel, New Office screens, Valiant chairs, Resolution X, AVE, Videocraft - $14,295 ex GST per Xero cash P&L.'
cf['E30']='=ROUND(18837.31*1.1,2)'
cf['O30']='Payables tab (Appspace, Spacera, Eptura, LiveU) plus the unbilled estimate. Week 2 to 12 Oct: Eptura $3,627, LiveU data pack $1,485 and the rest of the $18,837 ex GST in Xero cash P&L.'
cf['E31']='=ROUND((1115.04+46.69)*1.1,2)'
# overheads
for r in (35,36,37,38,39,41,42,44,45,46,47): cf[f'E{r}']=0
cf['E40']='=804.11+136'
cf['O40']='Optus mobiles BPAY, Telstra Nighthawk. Optus internet is on the card. Week 2 actual: Optus mobiles $804.11 + Telstra $136, both paid 8 Oct.'
cf['E43']=245
cf['O43']='Ricoh printing DD, Posh plants, plus the estimate. Week 2 actual: Posh plants $245 (8 Oct).'
cf['E52']=0
cf['O59']='Type the real closing balance (cheque + savings) in at the end of each week. Week 2 - key it Wed 14 Oct, then set the flag to 1.'
# make the week 2 actual cells look like week 1 actual cells (same style)
for r in range(13,53):
    if cf[f'E{r}'].value is not None and cf[f'C{r}'].has_style:
        cf[f'E{r}']._style=copy.copy(cf[f'C{r}']._style)

# ---------------- 3. Payables: extend the list to row 75, add the new Xero bills ----------------
# move the four total rows 61-64 down to 77-80
for r in range(61,65):
    for c in range(1,16):
        src=pb.cell(r,c); dst=pb.cell(r+16,c)
        dst.value=src.value; dst._style=copy.copy(src._style)
        src.value=None
    pb.row_dimensions[r+16].height=pb.row_dimensions[r].height
def fix(f):
    if not isinstance(f,str): return f
    return re.sub(r'(\$?[DEFGI]\$?19:\$?[DEFGI]\$?)60\b', r'\g<1>75', f)
for ws in (cf,pb):
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value,str) and c.value.startswith('='):
                c.value=fix(c.value)
pb['A77']='TOTAL bills listed from Xero (open at 28 Sep $283,042.53, plus the bills raised since - added 12 Oct)'
# template rows that are now real Xero bills
pb['B44']='INV179820 - CORE Intelligent Managed Services October'; pb['C44']=D(2026,10,15); pb['D44']=13566; pb['G44']=D(2026,10,15); pb['I44']='Xero bill'; pb['J44']='DD on the due date'
pb['B45']='INV180546 - Managed Services Azure October'; pb['C45']=D(2026,10,14); pb['D45']=692.84; pb['G45']=D(2026,10,14); pb['I45']='Xero bill'; pb['J45']='DD on the due date'
pb['B46']='INV-0120 - Digital growth marketing October'; pb['I46']='Xero bill'; pb['J46']='On the due date'
pb['B48']='INV-0060 - Aware Super national support October'; pb['I48']='Xero bill'; pb['J48']='On the due date, 31 Oct'
pb['B53']='INV01049843 - EH subscription September'; pb['C53']=D(2026,10,7); pb['D53']=2248.84; pb['I53']='Xero bill'
pb['B54']='INV10498230926 - Goget September'; pb['C54']=D(2026,10,1); pb['D54']=629.52; pb['I54']='Xero bill'
new=[
 ('Brookvale Insurance Brokers','INV1918462 - IT liability insurance (annual)',D(2026,10,14),35304.46,'EFT','Insurance',D(2026,10,14),'Xero bill','Annual renewal, GC approved. Due Wed 14 Oct - all six Brookvale bills go in one payment.'),
 ('Brookvale Insurance Brokers','INV1918336 - Removable / electronic equipment insurance (annual)',D(2026,10,14),11283.16,'EFT','Insurance',D(2026,10,14),'Xero bill','Annual renewal, GC approved.'),
 ('Brookvale Insurance Brokers','INV1918436 - Mecon contract works insurance (annual)',D(2026,10,14),8020.39,'EFT','Insurance',D(2026,10,14),'Xero bill','Annual renewal, GC approved.'),
 ('Brookvale Insurance Brokers','INV1920058 - Management liability insurance (annual)',D(2026,10,14),6778.21,'EFT','Insurance',D(2026,10,14),'Xero bill','Annual renewal, GC approved.'),
 ('Brookvale Insurance Brokers','INV1920113 - Business pack insurance (annual)',D(2026,10,14),3510.52,'EFT','Insurance',D(2026,10,14),'Xero bill','Annual renewal, GC approved.'),
 ('Brookvale Insurance Brokers','INV1914759 - Corporate travel insurance (annual)',D(2026,10,14),1342.24,'EFT','Insurance',D(2026,10,14),'Xero bill','Annual renewal, GC approved.'),
 ('Contact Group','INV457215 - DTTL Hobart AV fitout equipment & labour',D(2026,11,1),9652.54,'EFT','Equipment purchases',D(2026,11,2),'Xero bill','Due Sun 1 Nov, pay Mon 2 Nov (week 5).'),
 ('Contact Group','INV457214 - DTTL Hobart AV fitout equipment',D(2026,11,1),903.38,'EFT','Equipment purchases',D(2026,11,2),'Xero bill','Due Sun 1 Nov, pay Mon 2 Nov (week 5).'),
 ('Luminous Reputation Management','INV7200 - CTS 20 logo anniversary, balance of project fees',D(2026,10,30),4823.5,'EFT','Advertising',D(2026,10,30),'Xero bill','DL approved. On the due date (week 5).'),
 ('Easton Belle','INV00004009 - legal fees, general commercial file',D(2026,10,30),3932.5,'EFT','Professional fees',D(2026,10,30),'Xero bill','DL approved. On the due date (week 5).'),
 ('Ricoh Australia','INV15775235 - printing Sep',D(2026,10,31),210.2,'DD','Office',D(2026,10,30),'Xero bill','Direct debit'),
 ('Optus','INV000599632521 - Lugstein home internet Oct',D(2026,10,20),101.2,'BPAY','Telephone & internet',D(2026,10,20),'Xero bill','BPAY on the due date'),
 ('Flatworld Solutions','INV2026/USA-09032 - Seek data entry Sep',D(2026,10,6),92.21,'Card','Professional fees',D(2026,10,6),'Xero bill','PL card'),
]
r0=61
for i,row in enumerate(new):
    r=r0+i
    for c in range(1,16):
        pb.cell(r,c)._style=copy.copy(pb.cell(60,c)._style)
    for c,v in zip('ABCDEFGIJ',row):
        pb[f'{c}{r}']=v
    pb[f'H{r}']=f'=IF(G{r}<Assumptions!$B$5,"Before",IF(G{r}>Assumptions!$C$48,"After",INT((G{r}-Assumptions!$B$5)/7)+1))'
# insurance line now reads the Payables tab; the monthly estimate drops to nil so it isn't counted twice
tmpl=cf['B35'].value
for col,wk in zip('BDFHJ',range(44,49)):
    f=tmpl.replace('$B$44',f'$B${wk}').replace('$C$44',f'$C${wk}').replace('"Rent"','"Insurance"')
    cf[f'{col}42']=f+'+Assumptions!$B$34/5'
a['B34']=0
a['C34']="Was $3k a month as an estimate. The annual Brookvale renewals ($66,239, due 14 Oct) are now on the Payables tab under Insurance, so this is nil - put a number back in if something else comes up (icare, Asteron is quarterly and was paid 25 Sep)."
cf['O42']='Brookvale annual renewals $66,239 due Wed 14 Oct (six bills, GC approved) - Payables tab. Was a $3k a month estimate, now nil on top.'
pb['A2']=pb['A2'].value+' Updated 12 Oct: the bills raised since 28 Sep (Brookvale insurance renewals, Contact Group, Luminous, Easton Belle, First Focus and RJW October) are added at the bottom of the list.'

wb.calculation.fullCalcOnLoad=True
wb.save(OUT)
print('saved',OUT)
