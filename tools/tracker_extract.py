import openpyxl, json, re, html, datetime
wb = openpyxl.load_workbook('wb.xlsm', data_only=True)
ws = wb['Tracker']
def v(r, c):
    x = ws.cell(r, c).value
    if isinstance(x, str) and x.strip() == '': return None
    return x

# storage block starts
SUP, PRO, CON, WIP = 69, 97, 125, 153
recs = []

# ---- Support: BQ Date, BR Client, BS InvNo, BT Job, BU Desc, BV Amount, BW Notes
for r in range(12, 212):
    if not any(v(r, SUP + j) is not None for j in range(7)): continue
    recs.append(dict(source='Support', date=v(r,SUP), dept='Support', status=None,
        client=v(r,SUP+1), job=v(r,SUP+3), desc=v(r,SUP+4), invno=v(r,SUP+2),
        invval=v(r,SUP+5), wip=None, notes=v(r,SUP+6), _src_row=r))

# ---- Production
for r in range(12, 212):
    if not any(v(r, PRO + j) is not None for j in range(26)): continue
    recs.append(dict(source='Production', date=v(r,PRO), dept=v(r,PRO+1), status=None,
        client=v(r,PRO+2), job=v(r,PRO+4), desc=v(r,PRO+5), invno=v(r,PRO+12),
        invval=v(r,PRO+7), wip=None, notes=None, _src_row=r,
        email=v(r,PRO+3), eventdate=v(r,PRO+6), zoho=v(r,PRO+8), closed=v(r,PRO+9),
        curnum=v(r,PRO+10), posted=v(r,PRO+11), vfilm=v(r,PRO+13), vedit=v(r,PRO+14),
        pm=v(r,PRO+15), vpm=v(r,PRO+16), hours=v(r,PRO+17), vtotal=v(r,PRO+18),
        disc=v(r,PRO+19), crosshire=v(r,PRO+20), labint=v(r,PRO+21)))

# ---- Consulting
for r in range(12, 212):
    if not any(v(r, CON + j) is not None for j in range(18)): continue
    recs.append(dict(source='Consulting', date=v(r,CON), dept=v(r,CON+2), status=v(r,CON+1),
        client=v(r,CON+3), job=v(r,CON+4), desc=v(r,CON+5), invno=v(r,CON+8),
        invval=v(r,CON+9), wip=None, notes=v(r,CON+7), _src_row=r,
        qwilr=v(r,CON+6), labrev=v(r,CON+10), eqprev=v(r,CON+11), subrev=v(r,CON+12),
        labext=v(r,CON+13), eqpint=v(r,CON+14), subexp=v(r,CON+15)))

# ---- WIP
for r in range(12, 212):
    if not any(v(r, WIP + j) is not None for j in range(6)): continue
    recs.append(dict(source='WIP', date=v(r,WIP), dept=v(r,WIP+2), status=None,
        client=v(r,WIP+3), job=v(r,WIP+1), desc=v(r,WIP+4), invno=None,
        invval=None, wip=v(r,WIP+5), notes=None, _src_row=r))

# ---- comments (both kinds) with authors
people = {}
try:
    px = open('wb/xl/persons/person.xml').read()
    for m in re.finditer(r'<person[^>]*displayName="([^"]*)"[^>]*id="\{?([^"\}]*)', px):
        people[m.group(2).lower()] = html.unescape(m.group(1))
except Exception: pass
comments = []
tx = open('wb/xl/threadedComments/threadedComment1.xml').read()
for m in re.finditer(r'<threadedComment([^>]*)>(.*?)</threadedComment>', tx, re.S):
    at = m.group(1)
    ref = re.search(r'ref="([^"]+)"', at)
    dt  = re.search(r'dT="([^"]+)"', at)
    pid = re.search(r'personId="\{?([^"\}]*)', at)
    txt = ''.join(re.findall(r'<text>(.*?)</text>', m.group(2), re.S))
    comments.append(dict(kind='Threaded', ref=ref.group(1) if ref else '',
        when=(dt.group(1)[:10] if dt else ''),
        who=people.get(pid.group(1).lower(), '') if pid else '',
        text=html.unescape(txt).strip()))
cx = open('wb/xl/comments1.xml').read()
authors = [html.unescape(a) for a in re.findall(r'<author>(.*?)</author>', cx, re.S)]
for m in re.finditer(r'<comment ref="([^"]+)"[^>]*authorId="(\d+)"[^>]*>(.*?)</comment>', cx, re.S):
    txt = ''.join(re.findall(r'<t[^>]*>(.*?)</t>', m.group(3), re.S))
    comments.append(dict(kind='Note', ref=m.group(1), when='',
        who=authors[int(m.group(2))] if int(m.group(2)) < len(authors) else '',
        text=html.unescape(txt).strip()))

def enc(o):
    if isinstance(o, (datetime.datetime, datetime.date)): return {'__d': o.isoformat()}
    raise TypeError(str(type(o)))
json.dump({'records': recs, 'comments': comments}, open('extract.json','w'), default=enc)
print('records:', len(recs), {s: sum(1 for x in recs if x['source']==s) for s in ('Support','Production','Consulting','WIP')})
print('comments:', len(comments))
