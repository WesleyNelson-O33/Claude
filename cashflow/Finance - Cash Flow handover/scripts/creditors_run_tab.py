import openpyxl, datetime, copy, re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
D=datetime.datetime
wb=openpyxl.load_workbook('out/stage4_in.xlsx')
cf=wb['Cash Flow']; pb=wb['Payables']; a=wb['Assumptions']
assert all(pb.cell(r,11).value is None for r in range(1,81)), 'col K in use'

# ---------- new sheet ----------
cr=wb.create_sheet('Creditors Run', index=wb.sheetnames.index('Payables')+1)
title_font=copy.copy(cf['A1'].font); hdr_fill=copy.copy(pb['A18'].fill); hdr_font=copy.copy(pb['A18'].font)
blue=Font(name=hdr_font.name, size=hdr_font.size, color='FF0000FF')
cr['A1']='Creditors run - paste the Xero aged payables detail straight in'; cr['A1'].font=title_font
cr['A2']=("How it works: run the aged payables detail in Xero, export it, and paste the lines in from row 11 (columns A to K, same order as the report). "
          "Anything with a date in column A and a number in Gross gets counted - the contact and total lines the report adds are ignored. "
          "Every bill goes out on the run date in B4 unless you type a date in L, and an x in M leaves a line out (card bills - they sit in the card line). "
          "With B5 at 1 the Cash Flow Creditors run row takes these numbers and the Payables tab stops counting its Xero bill rows, so nothing is counted twice.")
cr['A2'].alignment=Alignment(wrap_text=True, vertical='top'); cr.merge_cells('A2:N2'); cr.row_dimensions[2].height=62
cr['A4']='Creditors run pay date'; cr['B4']=D(2026,11,2); cr['B4'].number_format='d mmm yy'; cr['B4'].font=blue
cr['C4']='Mon 2 Nov (1 Nov is a Sunday). Change it each month. Anything with its own date in column L ignores this.'
cr['A5']='Use this paste for the forecast? (1 = yes, 0 = use the Payables tab rows instead)'; cr['B5']=1; cr['B5'].font=blue
cr['A6']='Total open per the paste (what the forecast uses, before any x)'; cr['B6']='=SUMPRODUCT((A11:A510<>"")*ISNUMBER(I11:I510),I11:I510)'
cr['A7']='Report total per Xero (type it from the bottom of the report)'; cr['B7']=None; cr['B7'].font=blue
cr['A8']='Difference - should be nil'; cr['B8']='=IF(B7="","",B6-B7)'
for r in (6,7,8): cr[f'B{r}'].number_format=pb['D60'].number_format
heads=['Invoice Date','Source','Reference','Item Code','Description','Quantity','Unit Price (ex)','GST','Gross','Invoice Total','Status','Pay on (optional)','Exclude (x)','Pay date used']
for i,h in enumerate(heads,1):
    c=cr.cell(10,i,h); c.fill=copy.copy(hdr_fill); c.font=copy.copy(hdr_font); c.alignment=Alignment(wrap_text=True,vertical='center')
cr.row_dimensions[10].height=30
for r in range(11,511):
    cr[f'N{r}']=f'=IF(L{r}="",$B$4,L{r})'; cr[f'N{r}'].number_format='d mmm yy'
    cr[f'A{r}'].number_format='d mmm yy'; cr[f'L{r}'].number_format='d mmm yy'; cr[f'L{r}'].font=blue
    for col in 'GHIJ': cr[f'{col}{r}'].number_format=pb['D60'].number_format
for col,w in zip('ABCDEFGHIJKLMN',(11,9,16,9,52,8,13,11,12,12,16,13,10,12)): cr.column_dimensions[col].width=w
cr.freeze_panes='A11'
# prefill with what is open in Xero at 12 Oct
bills=[
 ('2026-09-18','INV950936412','Equipment purchases for 26052701 Aware SYD 28.06 Equipment [PO-00016916]',6457.51,645.75,7103.26,D(2026,10,19),''),
 ('2026-09-24','INV00133660','Annual Appspace licensing subscription (12 months) for 26091401 Aware Super [PO-00016941]',19077.66,1907.77,20985.43,D(2026,10,23),''),
 ('2026-09-25','INV15775235','Ricoh printing services September 2026 [paid via DD]',191.09,19.11,210.2,D(2026,10,30),''),
 ('2026-09-30','INV1918336','Brookvale - annual removable / electronic equipment insurance [GC App]',10300.62,982.54,11283.16,D(2026,10,14),''),
 ('2026-09-30','INV1920113','Brookvale - annual business pack insurance [GC App]',3211.96,298.57,3510.52,D(2026,10,14),''),
 ('2026-09-30','INV1914759','Brookvale - annual corporate travel insurance [GC App]',1246.49,95.75,1342.24,D(2026,10,14),''),
 ('2026-09-30','INV1918436','Brookvale - Mecon annual contract works insurance [GC App]',7338.73,681.66,8020.39,D(2026,10,14),''),
 ('2026-09-30','INV1920058','Brookvale - annual management liability insurance [GC App]',6195.64,582.57,6778.21,D(2026,10,14),''),
 ('2026-09-30','INV1918462','Brookvale - annual IT liability insurance [GC App]',32254.45,3050.02,35304.46,D(2026,10,14),''),
 ('2026-10-01','INV-0060','Spacera - Aware Super national support agreement October [PO-00016885]',5208.0,520.8,5728.8,D(2026,10,30),''),
 ('2026-09-30','INV01049843','Employment Hero subs renewal September 2026 [paid via PL CC]',2044.4,204.44,2248.84,None,'x'),
 ('2026-09-30','INV7200','Luminous - CTS 20 logo anniversary, balance of project fees [DL App]',4385.0,438.5,4823.5,D(2026,10,30),''),
 ('2026-09-30','INV10498230926','Goget September 2026 [paid via PL CC]',572.29,57.23,629.52,None,'x'),
 ('2026-10-01','INV179820','First Focus - CORE intelligent managed services October 2026 [paid via DD]',12332.73,1233.29,13566.0,D(2026,10,15),''),
 ('2026-10-06','INV000599632521','Optus - Lugstein home internet October 2026 [BPAY]',92.0,9.2,101.2,D(2026,10,20),''),
 ('2026-10-07','INV-0120','RJW - digital growth marketing October 2026',4250.0,425.0,4675.0,D(2026,10,21),''),
 ('2026-09-30','INV457214','Contact Group - equipment purchases for 2605605 DTTL Hobart AV fitout [PO-00016935]',821.25,82.13,903.38,D(2026,11,2),''),
 ('2026-09-30','INV457215','Contact Group - equipment & labour for 2605605 DTTL Hobart AV fitout [PO-00016921]',8775.0,877.54,9652.54,D(2026,11,2),''),
 ('2026-10-01','INV180546','First Focus - managed services Azure October 2026 [paid via DD]',629.85,62.99,692.84,D(2026,10,14),''),
 ('2026-10-06','INV2026/USA-09032','Flatworld - Seek data entry September 2026 [paid via PL CC]',92.21,0.0,92.21,None,'x'),
 ('2026-09-30','INV00004009','Easton Belle - legal fees, general commercial file [DL App]',3575.0,357.5,3932.5,D(2026,10,30),''),
 ('2026-09-24','INV260412412804','Energy Australia - electricity 23 Aug to 21 Sep [DD via PL CC]',632.95,63.3,696.25,None,'x'),
 ('2026-09-10','INV-18315','Persona Health - fitness for duty assessment (GC to confirm)',906.0,90.6,996.6,D(2026,10,21),''),
 ('2026-09-06','41 of 60','Employsure September 2026',800.0,80.0,880.0,D(2026,10,14),''),
]
listed=sum(b[5] for b in bills); asic=round(144499.05-listed,2)
bills.append(('2026-10-01','ASIC','ASIC - balance of the aged payables report at 12 Oct (check the actual bill)',asic,0.0,asic,D(2026,11,2),''))
for i,(dt,ref,desc,net,gst,tot,payon,exc) in enumerate(bills):
    r=11+i
    cr[f'A{r}']=D.fromisoformat(dt); cr[f'B{r}']='Bill'; cr[f'C{r}']=ref; cr[f'E{r}']=desc
    cr[f'G{r}']=net; cr[f'H{r}']=gst; cr[f'I{r}']=tot; cr[f'J{r}']=tot; cr[f'K{r}']='Awaiting Payment'
    if payon: cr[f'L{r}']=payon
    if exc: cr[f'M{r}']=exc
cr['A37']='Rows 11 to 35 are the open bills in Xero at 12 Oct ($144,499.05) typed in the report layout so the forecast has something to work from. Next month, clear them and paste the new report over the top.'
cr['A37'].font=Font(italic=True, size=9)

# ---------- Payables: flag Xero bill rows out of the category sums when the paste is on ----------
pb['K18']='Counted in the category lines?'; pb['K18']._style=copy.copy(pb['J18']._style)
for r in range(19,76):
    pb[f'K{r}']=f'=IF(AND(\'Creditors Run\'!$B$5=1,I{r}="Xero bill"),"No - in the creditors run","Yes")'
    pb[f'K{r}']._style=copy.copy(pb[f'I{r}']._style)
pb.column_dimensions['K'].width=22
# every category SUMIFS on the Cash Flow tab gets the extra condition
for row in cf.iter_rows(min_row=13,max_row=66):
    for c in row:
        v=c.value
        if isinstance(v,str) and v.startswith('=') and 'Payables!$D$19:$D$75' in v and 'SUMIFS' in v:
            c.value=re.sub(r'(SUMIFS\(Payables!\$D\$19:\$D\$75,)', r'\1Payables!$K$19:$K$75,"Yes",', v)

# ---------- Cash Flow row 48 ----------
cf['A48']='Creditors run (Xero aged payables paste)'
cond=("('Creditors Run'!$A$11:$A$510<>\"\")*ISNUMBER('Creditors Run'!$I$11:$I$510)*('Creditors Run'!$M$11:$M$510=\"\")"
      "*('Creditors Run'!$N$11:$N$510>=Assumptions!$B${w})*('Creditors Run'!$N$11:$N$510<=Assumptions!$C${w})")
for col,w in zip('BDFHJ',range(44,49)):
    cf[f'{col}48']=f"=IF('Creditors Run'!$B$5=1,SUMPRODUCT({cond.format(w=w)},'Creditors Run'!$I$11:$I$510),0)"
    cf[f'{col}48']._style=copy.copy(cf[f'{col}47']._style)
for col in 'EGIK': cf[f'{col}48']._style=copy.copy(cf[f'{col}47']._style)
cf['L48']='=SUM(B48,D48,F48,H48,J48)'; cf['M48']='=SUM(C48,E48,G48,I48,K48)'
cf['N48']='=(C48-B48)*C$8+(E48-D48)*E$8+(G48-F48)*G$8+(I48-H48)*I$8+(K48-J48)*K$8'
for col in 'LMN': cf[f'{col}48']._style=copy.copy(cf[f'{col}47']._style)
cf['A48']._style=copy.copy(cf['A47']._style)
cf['O48']="Creditors Run tab - paste the Xero aged payables detail in each month and it lands here on the run date (or the date in column L). Week 1 actual is the 1 Oct run. The Payables tab's Xero bill rows drop out of the lines above while this is switched on, so nothing is double counted."
cf['O48']._style=copy.copy(cf['O47']._style)
for col in 'BDEFGHIJKLMN':
    v=cf[f'{col}49'].value
    if isinstance(v,str) and f'{col}48' not in v: cf[f'{col}49']=v+f'+{col}48'
cf['A3']=cf['A3'].value+" The Creditors run row comes off the Creditors Run tab - paste the Xero aged payables detail in there each month."
wb.calculation.fullCalcOnLoad=True
wb.save('out/final.xlsx'); print('saved; ASIC balance', asic)
