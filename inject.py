"""Rebuild the HR copy: clean master + HR data + README + hidden columns."""
import sys
from openpyxl import load_workbook
from copy import copy
src_path, clean_path, out_path = sys.argv[1:4]
src=load_workbook(src_path); new=load_workbook(clean_path)
ss=src['Staff']; ns=new['Staff']
hs={ss.cell(row=4,column=c).value:ss.cell(row=4,column=c).column_letter for c in range(1,ss.max_column+1) if ss.cell(row=4,column=c).value}
hn={ns.cell(row=4,column=c).value:ns.cell(row=4,column=c).column_letter for c in range(1,ns.max_column+1) if ns.cell(row=4,column=c).value}
inputs=["Name","Surname","EmploymentType","DateOfBirth","StartDate","Casual Start","PTE Start","FTE Start","AnniversaryDate","PT Hrs/Week"]
for r in range(5,155):
    if ss[f'{hs["Name"]}{r}'].value is None: continue
    for h in inputs:
        v=ss[f'{hs[h]}{r}'].value
        if v is not None:
            c=ns[f'{hn[h]}{r}']; c.value=v
            if hasattr(v,'year'): c.number_format='dd/mm/yyyy'
for h,L in hs.items():
    if ss.column_dimensions[L].hidden and not h.startswith('h ') and h in hn: ns.column_dimensions[hn[h]].hidden=True
sl=src['Leave History']; nl=new['Leave History']
for r in range(1,sl.max_row+1):
    for c in range(1,13):
        v=sl.cell(row=r,column=c).value
        if v is not None: nl.cell(row=r,column=c).value=v
if 'READ ME' in src.sheetnames:
    sr=src['READ ME']; nr=new.create_sheet('READ ME',0)
    for row in sr.iter_rows():
        for c in row:
            if c.value is not None:
                d=nr.cell(row=c.row,column=c.column,value=c.value); d.font=copy(c.font); d.alignment=copy(c.alignment)
    for k,dim in sr.column_dimensions.items(): nr.column_dimensions[k].width=dim.width
    nr.sheet_properties.tabColor="1F4E78"
ms=new['Milestones']; key=ms['C5'].value
for r in range(5,155):
    if ns[f'A{r}'].value=='Isaac': ms['C5']=key[:-1]+str(r); break
# Settings values from the HR file
for cell in ('C5','C6','C7','C8','C9','C19','C31'):
    v=src['Settings'][cell].value
    if v is not None: new['Settings'][cell].value=v
new.calculation.fullCalcOnLoad=True; new.save(out_path); print('saved', out_path, new.sheetnames)
