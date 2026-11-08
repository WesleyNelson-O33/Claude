import openpyxl, pickle, sys, re, datetime as dt
wb=openpyxl.load_workbook(sys.argv[1],read_only=True,data_only=True)
ws=wb.active
js=[];cur=None
for r in ws.iter_rows(min_row=6,values_only=True):
    a=r[0]
    if isinstance(a,str) and a.startswith('ID '):
        cur={'title':a,'lines':[]};js.append(cur)
    elif isinstance(a,dt.datetime) and cur is not None:
        cur['lines'].append(dict(date=a,id=r[1],code=str(r[2]),acct=r[3],dr=r[4] or 0,cr=r[5] or 0,posted=r[6],by=r[7]))
pickle.dump(js,open(sys.argv[2],'wb'))
print(len(js))
