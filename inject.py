"""Rebuild the HR copy: clean master + HR data + README + hidden columns + HR formatting + tally preload."""
import sys, re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as L
from copy import copy
src_path, clean_path, out_path = sys.argv[1:4]
tally_path = sys.argv[4] if len(sys.argv) > 4 else None
src=load_workbook(src_path); new=load_workbook(clean_path)
def copy_style(a,b):
    b.font=copy(a.font); b.fill=copy(a.fill); b.border=copy(a.border); b.alignment=copy(a.alignment); b.number_format=a.number_format
ss=src['Staff']; ns=new['Staff']
hs={ss.cell(row=4,column=c).value:c for c in range(1,ss.max_column+1) if ss.cell(row=4,column=c).value}
hn={ns.cell(row=4,column=c).value:c for c in range(1,ns.max_column+1) if ns.cell(row=4,column=c).value}
inputs=["Name","Surname","EmploymentType","DateOfBirth","StartDate","Casual Start","PTE Start","FTE Start","AnniversaryDate","PT Hrs/Week"]
for r in range(5,155):
    if ss.cell(row=r,column=hs["Name"]).value is None: continue
    for h in inputs:
        v=ss.cell(row=r,column=hs[h]).value
        if v is not None:
            c=ns.cell(row=r,column=hn[h]); c.value=v
            if hasattr(v,'year'): c.number_format='dd/mm/yyyy'
# formatting by header
for h,cs in hs.items():
    if h in hn and not h.startswith('h '):
        cn=hn[h]
        for r in range(1,155): copy_style(ss.cell(row=r,column=cs), ns.cell(row=r,column=cn))
        ns.column_dimensions[L(cn)].width=ss.column_dimensions[L(cs)].width
        ns.column_dimensions[L(cn)].hidden=ss.column_dimensions[L(cs)].hidden
for h in ("Earned this month (days)","Milestone this month"):
    if h in hn and h not in hs:
        for r in range(3,155): copy_style(ss.cell(row=r,column=hs["Hours to Accrue"]), ns.cell(row=r,column=hn[h]))
        ns.cell(row=4,column=hn[h]).value=h
if "Milestone this month" in hn:
    for r in range(5,155): ns.cell(row=r,column=hn["Milestone this month"]).number_format='dd/mm/yyyy'
for r in range(1,5):
    ns.row_dimensions[r].height=ss.row_dimensions[r].height
    for c in range(1,4): copy_style(ss.cell(row=r,column=c), ns.cell(row=r,column=c))
ns.freeze_panes=ss.freeze_panes
# Leave History data A:L
sl=src['Leave History']; nl=new['Leave History']
for r in range(1,sl.max_row+1):
    for c in range(1,13):
        v=sl.cell(row=r,column=c).value
        if v is not None: nl.cell(row=r,column=c).value=v
for name in ('Settings','Leave History'):
    a=src[name]; b=new[name]
    for row in a.iter_rows(min_row=1,max_row=(60 if name=='Settings' else 1)):
        for c in row:
            if c.has_style: copy_style(c, b.cell(row=c.row,column=c.column))
    for k,dim in a.column_dimensions.items(): b.column_dimensions[k].width=dim.width
for cell in ('C32','C33'): new['Settings'][cell].number_format='dd/mm/yyyy'
# Settings values the HR team set
for cell in ('C5','C6','C7','C8','C9','C19','C31'):
    v=src['Settings'][cell].value
    if v is not None: new['Settings'][cell].value=v
# READ ME
if 'READ ME' in src.sheetnames:
    sr=src['READ ME']; nr=new.create_sheet('READ ME',0)
    for row in sr.iter_rows():
        for c in row:
            if c.value is not None:
                d=nr.cell(row=c.row,column=c.column,value=c.value); d.font=copy(c.font); d.alignment=copy(c.alignment)
    for k,dim in sr.column_dimensions.items(): nr.column_dimensions[k].width=dim.width
    nr.sheet_properties.tabColor="1F4E78"
for name in src.sheetnames:
    if name in new.sheetnames: new[name].sheet_properties.tabColor=src[name].sheet_properties.tabColor
# tally preload
if tally_path:
    names=["Christopher Lehmann","Daniel Sobkowski","Danielle Hurley","Duncan Lugstein","Jordan Sexty","Kyle Krishnappa","Milo Rankin","Patrick Lee","Simon Gruenefeld","Tyler Wood","Wade Tonna","Isaac Butterworth","Blake Crisford","Carlo Daru"]
    t=new['HR Tally Check']
    for p,nm in enumerate(names): t.cell(row=7,column=2+p).value=nm
    i=0
    for line in open(tally_path):
        parts=line.rstrip('\n').split('\t')
        if not re.match(r'[A-Z][a-z]{2}-\d{2}',parts[0]): continue
        t.cell(row=8+i,column=1).value=parts[0]
        for p in range(len(names)):
            v=parts[p+1].strip() if p+1<len(parts) else ''
            if v: t.cell(row=8+i,column=2+p).value=float(v)
        i+=1
# Milestones default = Isaac
ms=new['Milestones']; key=ms['C5'].value
for r in range(5,155):
    if ns[f'A{r}'].value=='Isaac': ms['C5']=key[:-1]+str(r); break
new.calculation.fullCalcOnLoad=True; new.save(out_path); print('saved', out_path, new.sheetnames)
