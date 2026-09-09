import json, datetime, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo, TableFormula
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as L

D = json.load(open('extract.json'))
def dec(o):
    if isinstance(o, dict) and '__d' in o: return datetime.datetime.fromisoformat(o['__d'])
    return o
recs = [{k: dec(v) for k, v in r.items()} for r in D['records']]

NAVY='3E5066'; LIGHT='5B708A'; CALCHDR='747474'; GREY='EDEDED'
YELLOW='FFF3C4'; RULE='B7B7B7'; F='Arial'
MONEY='_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
PCT='0.0%'; DATE='dd-mmm-yy'
thin=Side(style='thin', color=RULE); box=Border(left=thin,right=thin,top=thin,bottom=thin)
wb = openpyxl.Workbook()

# ---------------------------------------------------------------- LISTS
ls = wb.active; ls.title='Lists'
ls['A1']='Control values — edit here and every dropdown, check and report line follows'
ls['A1'].font=Font(F,11,bold=True,color=NAVY)
ls['A3']='Financial year start';        ls['B3']=datetime.datetime(2026,7,1)
ls['A4']='Financial year end (excl.)';  ls['B4']=datetime.datetime(2027,7,1)
for c in ('B3','B4'):
    ls[c].number_format=DATE; ls[c].font=Font(F,10,color='0000FF')
    ls[c].fill=PatternFill('solid',fgColor=YELLOW); ls[c].border=box
DEPTMAP=[('Consulting','Consulting'),('Integration','Consulting'),('Onsite','Production'),
         ('Other','Production'),('Production','Production'),('Support','Support'),('Video','Production')]
def listcol(col,title,items,row=3):
    ls[f'{col}{row}']=title
    ls[f'{col}{row}'].font=Font(F,10,bold=True,color='FFFFFF')
    ls[f'{col}{row}'].fill=PatternFill('solid',fgColor=NAVY)
    ls[f'{col}{row}'].alignment=Alignment(horizontal='center'); ls[f'{col}{row}'].border=box
    for i,x in enumerate(items):
        c=ls[f'{col}{row+1+i}']; c.value=x; c.font=Font(F,10); c.border=box
listcol('D','Type',['Invoice','WIP'])
listcol('F','Department',[d for d,_ in DEPTMAP])
listcol('G','Reports as',[r for _,r in DEPTMAP])
listcol('I','Status',['In Progress','Completed','Cancelled','Closed'])
listcol('K','Yes / No',['Y','N'])
ls['F12']='A record whose Type is WIP always reports as WIP, whatever its department.'
ls['F12'].font=Font(F,9,italic=True,color='555555')
for col,w in zip('ABDFGIK',[26,14,14,16,16,16,12]): ls.column_dimensions[col].width=w

# ---------------------------------------------------------------- DATA
ds = wb.create_sheet('Data')
HDRS=[('No.',6),('Type',10),('Date',11),('Department',13),('Reports as',12),('Status',12),
 ('Client',22),('Job Number',13),('Description',40),('Invoice Number',15),('Amount',14),
 ('Net Revenue',14),('Margin',13),('Margin %',10),('Date Check',13),('Duplicate Check',19),('Notes',34),
 ('Client Email',26),('Event Date',11),('Zoho Number',13),('Current Number',14),('Closed',9),
 ('Invoice Posted to Xero',13),('Video Filming',11),('Video Editing',11),('Project Management',12),
 ('Video Project Management',13),('Production Labour Hours',12),('Video Total',13),
 ('Discounts Included',13),('Discounts as % of Net',12),('Cross Hire Expense',14),
 ('Labour Expense (Internal)',14),('Qwilr Link',34),('Labour Revenue',14),('Equipment Revenue',14),
 ('Subscription Revenue',14),('Labour Expense (External)',14),('Equipment Expense (Internal)',14),
 ('Subscription & Licences Expense',14)]
CALC={1,5,12,13,14,15,16,31}
for i,(h,w) in enumerate(HDRS,1):
    c=ds.cell(1,i,h); c.font=Font(F,10,bold=True,color='FFFFFF')
    c.fill=PatternFill('solid',fgColor=(CALCHDR if i in CALC else LIGHT if 18<=i<=33 else NAVY))
    c.alignment=Alignment(horizontal='center',vertical='top',wrap_text=True); c.border=box
    ds.column_dimensions[L(i)].width=w
ds.row_dimensions[1].height=46

FORM={
 1 :'IF(COUNTA($B{r}:$K{r})=0,"",ROW()-1)',
 5 :'IF($B{r}="","",IF($B{r}="WIP","WIP",IFERROR(INDEX(Lists!$G$4:$G$10,MATCH($D{r},Lists!$F$4:$F$10,0)),"Unmapped")))',
 12:'IF($K{r}="","",IF($E{r}="Production",$K{r}-N($AD{r}),$K{r}))',
 13:'IF($E{r}="Production",IF($K{r}="","",$K{r}-(N($AF{r})+N($AG{r}))),'
    'IF($E{r}="Consulting",IF($K{r}="","",$K{r}-(N($AL{r})+N($AM{r})+N($AN{r}))),""))',
 14:'IFERROR($M{r}/$K{r},"")',
 15:'IF(COUNTA($B{r}:$K{r})=0,"",IF($C{r}="","No date",'
    'IF(AND($C{r}>=Lists!$B$3,$C{r}<Lists!$B$4),"In FY27","Outside FY27")))',
 16:'IF($E{r}<>"WIP","",IF($H{r}="","",'
    'IF(COUNTIFS($E$2:$E$20000,"<>WIP",$H$2:$H$20000,$H{r})>0,"Job already invoiced","")))',
 31:'IF($E{r}<>"Production","",IFERROR($AD{r}/$L{r},""))'}

KEYS=['','','date','dept','','status','client','job','desc','invno','amount','','','','','','notes',
 'email','eventdate','zoho','curnum','closed','posted','vfilm','vedit','pm','vpm','hours','vtotal',
 'disc','','crosshire','labint','qwilr','labrev','eqprev','subrev','labext','eqpint','subexp']
order={'Support':0,'Production':1,'Consulting':2,'WIP':3}
recs.sort(key=lambda r:(order[r['source']], r['_src_row']))
for rec in recs:
    rec['type']   = 'WIP' if rec['source']=='WIP' else 'Invoice'
    rec['amount'] = rec['wip'] if rec['source']=='WIP' else rec['invval']
    if rec['source']=='Support': rec['dept']='Support'
N=len(recs)
for i,rec in enumerate(recs):
    r=i+2
    ds.cell(r,2,rec['type'])
    for ci,key in enumerate(KEYS,1):
        if key and rec.get(key) is not None: ds.cell(r,ci,rec[key])
    for ci,f in FORM.items(): ds.cell(r,ci,'='+f.format(r=r))
LAST=N+1
MONEY_COLS=[11,12,13,29,30,32,33,35,36,37,38,39,40]; PCT_COLS=[14,31]; DATE_COLS=[3,19]
CTR=[1,2,3,5,8,10,15,16,19,20,21,22,23,24,25,26,27,28]
for r in range(2,LAST+1):
    ds.row_dimensions[r].height=15
    for ci in range(1,41):
        c=ds.cell(r,ci); c.font=Font(F,10); c.border=box
        c.alignment=Alignment(horizontal=('center' if ci in CTR else 'left'),vertical='center')
        if ci in MONEY_COLS: c.number_format=MONEY
        elif ci in PCT_COLS: c.number_format=PCT
        elif ci in DATE_COLS: c.number_format=DATE
        if ci in CALC: c.fill=PatternFill('solid',fgColor=GREY)

tbl=Table(displayName='tblRevenue', ref=f'A1:AN{LAST}')
tbl.tableStyleInfo=TableStyleInfo(name='TableStyleLight1',showRowStripes=False,showColumnStripes=False)
ds.add_table(tbl)
for rng,f1 in ((f'B2:B{LAST}','Lists!$D$4:$D$5'),(f'D2:D{LAST}','Lists!$F$4:$F$10'),
               (f'F2:F{LAST}','Lists!$I$4:$I$7'),(f'V2:W{LAST}','Lists!$K$4:$K$5')):
    dv=DataValidation(type='list',formula1=f1,allow_blank=True,showErrorMessage=True,showInputMessage=True)
    dv.error='Pick a value from the list on the Lists tab.'
    ds.add_data_validation(dv); dv.add(rng)
ds.column_dimensions.group('R','AG',outline_level=1,hidden=False)
ds.column_dimensions.group('AH','AN',outline_level=1,hidden=False)
ds.sheet_properties.outlinePr.summaryRight=False
ds.freeze_panes='F2'; ds.sheet_view.showGridLines=False
print('data rows', N)

# ---------------------------------------------------------------- DASHBOARD
dsh=wb.create_sheet('Dashboard'); dsh.sheet_view.showGridLines=False
def title(ws,cell,text,size=13):
    ws[cell]=text; ws[cell].font=Font(F,size,bold=True,color=NAVY)
def hdrow(ws,row,labels,fill=NAVY):
    for i,t in enumerate(labels):
        c=ws.cell(row,1+i,t); c.font=Font(F,10,bold=True,color='FFFFFF')
        c.fill=PatternFill('solid',fgColor=fill); c.border=box
        c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[row].height=30
title(dsh,'A1','FY 2026-27 — Monthly revenue by report line')
hdrow(dsh,2,['Month','Support','Production','Consulting','WIP','Total Revenue','P&L Total','Variance'])
for i in range(12):
    r=3+i; m=7+i; y=2026+(m-1)//12; m=(m-1)%12+1
    dsh.cell(r,1,datetime.datetime(y,m,1)).number_format='mmm-yy'
    for ci in range(2,6):
        dsh.cell(r,ci,f'=SUMIFS(Data!$L$2:$L$20000,Data!$E$2:$E$20000,{L(ci)}$2,'
                      f'Data!$C$2:$C$20000,">="&$A{r},Data!$C$2:$C$20000,"<"&EDATE($A{r},1))')
    dsh.cell(r,6,f'=SUM($B{r}:$E{r})')
    dsh.cell(r,8,f'=IF($G{r}="","",IF(ABS($F{r}-$G{r})<0.99,"Reconciled",$F{r}-$G{r}))')
TR=15
dsh.cell(TR,1,'Total')
for ci in range(2,8): dsh.cell(TR,ci,f'=SUM({L(ci)}3:{L(ci)}14)')
dsh.cell(TR,8,'=IF(COUNT($G$3:$G$14)=0,"",'
              'IF(ABS(SUMPRODUCT(($G$3:$G$14<>"")*$F$3:$F$14)-SUM($G$3:$G$14))<0.99,"Reconciled",'
              'SUMPRODUCT(($G$3:$G$14<>"")*$F$3:$F$14)-SUM($G$3:$G$14)))')
for r in range(3,TR+1):
    for ci in range(1,9):
        c=dsh.cell(r,ci); c.font=Font(F,10,bold=(r==TR)); c.border=box
        c.alignment=Alignment(horizontal='center' if ci in (1,8) else 'right',vertical='center')
        if 2<=ci<=7: c.number_format=MONEY
        if ci==7 and r<TR: c.fill=PatternFill('solid',fgColor=YELLOW); c.font=Font(F,10,color='0000FF')
        elif r==TR: c.fill=PatternFill('solid',fgColor=GREY)
dsh['A17']='P&L Total (yellow) is the only figure you type here. Variance reads "Reconciled" under $0.99.'
dsh['A18']='The Total row measures only the months where a P&L Total has been entered, so a part-year shows no false gap.'
for c in ('A17','A18'): dsh[c].font=Font(F,9,italic=True,color='555555')
title(dsh,'A21','Reconciliation — is every record reaching the monthly view?')
hdrow(dsh,22,['Report line','Total in tracker','In FY27 months','No date','Dated outside FY27','Adds up?'])
for i,s in enumerate(['Support','Production','Consulting','WIP']):
    r=23+i; dsh.cell(r,1,s)
    dsh.cell(r,2,f'=SUMIF(Data!$E$2:$E$20000,$A{r},Data!$L$2:$L$20000)')
    for ci,tag in ((3,'In FY27'),(4,'No date'),(5,'Outside FY27')):
        dsh.cell(r,ci,f'=SUMIFS(Data!$L$2:$L$20000,Data!$E$2:$E$20000,$A{r},Data!$O$2:$O$20000,"{tag}")')
    dsh.cell(r,6,f'=IF(ABS($B{r}-SUM($C{r}:$E{r}))<0.01,"OK","Check")')
dsh.cell(27,1,'Total')
for ci in range(2,6): dsh.cell(27,ci,f'=SUM({L(ci)}23:{L(ci)}26)')
dsh.cell(27,6,'=IF(ABS($C$27-$F$15)<0.01,"Matches monthly grid","Does not match monthly grid")')
for r in range(23,28):
    for ci in range(1,7):
        c=dsh.cell(r,ci); c.font=Font(F,10,bold=(r==27)); c.border=box
        c.alignment=Alignment(horizontal='center' if ci in (1,6) else 'right',vertical='center')
        if 2<=ci<=5: c.number_format=MONEY
        if r==27: c.fill=PatternFill('solid',fgColor=GREY)
dsh['A29']=('"No date" and "Dated outside FY27" are different problems. A blank date needs filling in. '
            'A date outside FY27 is a real date from another financial year.')
dsh['A29'].font=Font(F,9,italic=True,color='555555')
for col,w in zip('ABCDEFGH',[14,17,17,17,17,17,15,24]): dsh.column_dimensions[col].width=w
dsh.freeze_panes='A3'

# ---------------------------------------------------------------- CHECKS
ck=wb.create_sheet('Checks'); ck.sheet_view.showGridLines=False
title(ck,'A1','Data quality checks')
ck['A2']='Live formulas — they update the moment you change anything on the Data tab.'
ck['A2'].font=Font(F,9,italic=True,color='555555')
hdrow(ck,4,['#','What is being checked','Records','Value','Status','What to do'])
CHECKS=[
 ('Records with no date','=COUNTIF(Data!$O$2:$O$20000,"No date")',
  '=SUMIF(Data!$O$2:$O$20000,"No date",Data!$L$2:$L$20000)',
  'Filter Date Check to "No date" on the Data tab and fill them in. This money is in no month until you do.'),
 ('Records dated outside FY 2026-27','=COUNTIF(Data!$O$2:$O$20000,"Outside FY27")',
  '=SUMIF(Data!$O$2:$O$20000,"Outside FY27",Data!$L$2:$L$20000)',
  'Real dates from another financial year. Decide whether they belong in this tracker at all.'),
 ('WIP jobs that have also been invoiced','=COUNTIF(Data!$P$2:$P$20000,"Job already invoiced")',
  '=SUMIF(Data!$P$2:$P$20000,"Job already invoiced",Data!$L$2:$L$20000)',
  'Filter Duplicate Check on the Data tab. Each needs a reversing WIP entry or it is double counted.'),
 ('Departments that do not map to a report line','=COUNTIF(Data!$E$2:$E$20000,"Unmapped")','',
  'The department is not on the Lists tab. Correct the record, or add the department and its report line.'),
 ('Records with a Type but no Department',
  '=SUMPRODUCT((Data!$B$2:$B$20000<>"")*(Data!$D$2:$D$20000=""))','',
  'Every record needs a department so it lands on the right report line.'),
 ('Records with money but no client',
  '=SUMPRODUCT((Data!$K$2:$K$20000<>0)*(Data!$K$2:$K$20000<>"")*(Data!$G$2:$G$20000=""))','',
  'A revenue line with no client cannot be traced back to anything.'),
 ('Records with money but no job number',
  '=SUMPRODUCT((Data!$K$2:$K$20000<>0)*(Data!$K$2:$K$20000<>"")*(Data!$H$2:$H$20000=""))','',
  'Without a job number the WIP duplicate check cannot see this record.'),
 ('Monthly grid vs total in tracker','','=Dashboard!$F$15-Dashboard!$C$27',
  'Should be zero. Anything else means the Dashboard and the Data tab disagree.')]
for i,(what,cnt,val,todo) in enumerate(CHECKS):
    r=5+i
    ck.cell(r,1,i+1); ck.cell(r,2,what)
    if cnt: ck.cell(r,3,cnt)
    if val: ck.cell(r,4,val).number_format=MONEY
    ck.cell(r,5,f'=IF(AND(N($C{r})=0,ROUND(N($D{r}),2)=0),"OK","Review")')
    ck.cell(r,6,todo)
    for ci in range(1,7):
        c=ck.cell(r,ci); c.font=Font(F,10); c.border=box
        c.alignment=Alignment(horizontal=('center' if ci in (1,3,5) else 'right' if ci==4 else 'left'),
                              vertical='center',wrap_text=(ci==6))
    ck.row_dimensions[r].height=30
for col,w in zip('ABCDEF',[5,44,10,16,11,74]): ck.column_dimensions[col].width=w
dupe={}
for rec in recs:
    if rec['type']!='WIP' and rec.get('job'): dupe.setdefault(str(rec['job']),[]).append(rec)
rows=[(str(r['job']),r.get('client'),r.get('amount'),o['dept'],o.get('amount'))
      for r in recs if r['type']=='WIP' and r.get('job') and str(r['job']) in dupe
      for o in dupe[str(r['job'])]]
R0=5+len(CHECKS)+2
title(ck,f'A{R0}','WIP jobs that already appear on an invoiced record')
ck.cell(R0+1,1,'Snapshot from the rebuild. The live version is the Duplicate Check column on the Data tab.')\
  .font=Font(F,9,italic=True,color='555555')
hdrow(ck,R0+2,['Job Number','Client','WIP Amount','Invoiced dept','Invoiced Amount'])
for i,row in enumerate(rows):
    r=R0+3+i
    for ci,x in enumerate(row,1):
        c=ck.cell(r,ci,x); c.font=Font(F,10); c.border=box
        c.alignment=Alignment(horizontal='center' if ci in (1,4) else 'left' if ci==2 else 'right')
        if ci in (3,5): c.number_format=MONEY

# ---------------------------------------------------------------- LEGACY NOTES
ln=wb.create_sheet('Legacy Notes'); ln.sheet_view.showGridLines=False
title(ln,'A1','Notes and comments recovered from the old workbook')
for i,t in enumerate([
 'In the old file these were Excel comments stuck to cell positions, not to jobs. The macro swapped the data',
 'underneath them every time you changed department, so they had already drifted off the records they describe.',
 'They are listed here with their original cell reference. Paste each into the Notes column on the Data tab',
 'against the right record. Nothing has been guessed — no note has been assigned to a record for you.']):
    ln.cell(2+i,1,t).font=Font(F,9,italic=True,color='555555')
hdrow(ln,7,['Original cell','Type','Author','Date','Note text'])
cm=sorted(D['comments'],key=lambda c:(int(''.join(ch for ch in c['ref'] if ch.isdigit()) or 0),
                                      ''.join(ch for ch in c['ref'] if ch.isalpha())))
w=0
for c in cm:
    if not c['text']: continue
    r=8+w; w+=1
    for ci,x in enumerate([c['ref'],c['kind'],c['who'],c['when'],c['text']],1):
        cell=ln.cell(r,ci,x); cell.font=Font(F,10); cell.border=box
        cell.alignment=Alignment(horizontal='center' if ci in (1,2,4) else 'left',
                                 vertical='top',wrap_text=(ci==5))
    ln.row_dimensions[r].height=14*max(1,c['text'].count('\n')+1)
for col,w2 in zip('ABCDE',[13,11,22,12,96]): ln.column_dimensions[col].width=w2
ln.freeze_panes='A8'

# ---------------------------------------------------------------- READ ME
rm=wb.create_sheet('Read Me',0); rm.sheet_view.showGridLines=False
rm.column_dimensions['A'].width=3; rm.column_dimensions['B'].width=112
def para(row,text,style='body'):
    c=rm.cell(row,2,text)
    c.font={'h1':Font(F,16,bold=True,color=NAVY),'h2':Font(F,11,bold=True,color=NAVY),
            'note':Font(F,9,italic=True,color='555555')}.get(style,Font(F,10))
    c.alignment=Alignment(vertical='top',wrap_text=True)
    rm.row_dimensions[row].height=26 if style=='h1' else 13
    return row+1
CONTENT=[
 ('h1','FY 2026-27 Revenue Tracker'),
 ('note','Rebuilt from FY27_Revenue_Tracker_v2.xlsm. Your original file has not been changed.'),
 ('body',''),
 ('h2','How a record works'),
 ('body','Every record needs two things: a Type and a Department. Between them they decide which Dashboard'),
 ('body','column the money lands in.'),
 ('body',''),
 ('body','    Type            Invoice or WIP. Pick WIP for a work-in-progress movement, Invoice for everything else.'),
 ('body','    Department      Consulting, Integration, Onsite, Other, Production, Support or Video.'),
 ('body','    Amount          One column, always. For a WIP record this is the movement, and it can be negative.'),
 ('body',''),
 ('body','"Reports as" is calculated for you and shows which Dashboard column the record feeds. Type WIP always'),
 ('body','reports as WIP whatever the department. Otherwise the department decides: Support goes to Support,'),
 ('body','Consulting and Integration go to Consulting, and Production, Video, Other and Onsite go to Production.'),
 ('body','That mapping is on the Lists tab if you ever need to change it.'),
 ('body',''),
 ('h2','Typing a new record'),
 ('body','Click the first empty row under the table and start typing. The table extends itself and every grey'),
 ('body','column fills in automatically. Grey columns are calculated — do not type in them.'),
 ('body',''),
 ('body','    Reports as      Which Dashboard column this record feeds.'),
 ('body','    Net Revenue     What actually rolls up. Production deducts Discounts Included; everything else is Amount.'),
 ('body','    Margin          Production: Amount less Cross Hire and Internal Labour. Consulting: Amount less the'),
 ('body','                    three expense columns.'),
 ('body','    Date Check      Flags "No date" or "Outside FY27" so nothing goes missing silently.'),
 ('body','    Duplicate Check Flags a WIP job that has already been invoiced somewhere else.'),
 ('body',''),
 ('body','To see one department or one report line, click the filter arrow in its header. The Production and'),
 ('body','Consulting detail columns are grouped — use the + and - buttons above the column letters to collapse them.'),
 ('body',''),
 ('h2','The tabs'),
 ('body','Data            Every record, one row each. 181 brought across from the old file.'),
 ('body','Dashboard       Revenue by month and report line, plus the reconciliation. Type the P&L Total; nothing else.'),
 ('body','Checks          Live data quality checks. Look here first each month.'),
 ('body','Legacy Notes    The comments recovered from the old file, for you to reattach.'),
 ('body','Lists           Dropdown values, the department mapping, and the financial year dates.'),
 ('body',''),
 ('h2','Carried across, still needing a decision'),
 ('body','None of these have been changed. They are flagged so you can decide.'),
 ('body',''),
 ('body','1.  29 records have no date, so they appear in no month. See Checks.'),
 ('body','2.  16 Consulting records carry FY26 dates, May 2025 to June 2026. The old file lumped these in with the'),
 ('body','    undated ones and called them all "Undated", which hid the problem. They are shown separately now.'),
 ('body','3.  Five WIP jobs have also been invoiced. One is a clean reversal, the rest are not. See Checks.'),
 ('body','4.  One record has the department "Onsite", which was on no list in the old file. It has been kept and'),
 ('body','    added to the Lists tab so nothing was lost. Reclassify it if it should be something else.'),
 ('body','5.  Margin on Production is calculated from Amount, not from Net Revenue, so discounts are not deducted'),
 ('body','    before margin. That is how the old file did it. Worth checking it is what you want.'),
 ('body',''),
 ('h2','One thing to know about the numbers'),
 ('body','The old file was showing stale Production figures, about $26,792 higher than its own formulas produced.'),
 ('body','Everything here is calculated fresh, so July and August Production read lower. The lower ones are correct.')]
row=2
for style,text in CONTENT: row=para(row,text,style)
for ws in wb.worksheets: ws.sheet_properties.tabColor=NAVY
OUT='FY27_Revenue_Tracker_v4.xlsx'; wb.save(OUT)

# Excel only auto-fills a formula into a new table row when the table part declares
# the column as calculated. openpyxl builds tableColumns at save time, so this is
# injected afterwards, keyed on the header name.
import zipfile, shutil, re as _re
from xml.sax.saxutils import escape as _esc
NAMES={HDRS[i-1][0]: FORM[i].format(r=2) for i in FORM}
src=zipfile.ZipFile(OUT); parts={n: src.read(n) for n in src.namelist()}; src.close()
tpath=[n for n in parts if _re.match(r'xl/tables/table\d+\.xml$', n)][0]
xml=parts[tpath].decode('utf-8'); hits=0
def _inject(m):
    global hits
    whole, name = m.group(0), m.group(1)
    if name not in NAMES or 'calculatedColumnFormula' in whole: return whole
    hits += 1
    body='<calculatedColumnFormula>'+_esc(NAMES[name])+'</calculatedColumnFormula>'
    return whole[:-2]+'>'+body+'</tableColumn>' if whole.endswith('/>') else whole+body
xml=_re.sub(r'<tableColumn\b[^>]*name="([^"]+)"[^>]*(?:/>|>)', _inject, xml)
parts[tpath]=xml.encode('utf-8')
with zipfile.ZipFile(OUT,'w',zipfile.ZIP_DEFLATED) as z:
    for n,b in parts.items(): z.writestr(n,b)
assert hits==len(NAMES), f'injected {hits} of {len(NAMES)} calculated columns'
print('saved',OUT,'| rows',N,'| notes',w,'| auto-fill columns',hits)
