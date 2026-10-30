import csv,re,json,datetime as dt
F='/tmp/claude-0/-home-user-notion-goapl/3894af6e-7006-5c78-8a81-d4969e11e621/scratchpad/fcc/'
hd={}
for l in open(F+'HD.dat',encoding='latin-1'):
    p=l.rstrip('\n').split('|')
    if len(p)<45: continue
    hd[p[1]]=(p[5],p[6],p[7],p[8],p[43] if len(p)>43 else '')  # status, svc, grant, expire, effective/last action
def d(s):
    try: return dt.datetime.strptime(s,'%m/%d/%Y')
    except: return None
kind=re.compile(r'\b(CITY|VILLAGE|TOWN|TOWNSHIP|COUNTY|TWP|CNTY)\b',re.I)
bad=re.compile(r'SCHOOL|DISTRICT|DIST\b|PARK DIST|COLLEGE|UNIVERSITY|HOSPITAL|LIBRARY|HOUSING|FIRE PROTECTION|FPD|SANITARY|AIRPORT|AUTHORITY|FOREST PRES|CONSERVATION|MOSQUITO|WATER RECLAMATION|SHERIFF|STATE OF|BOARD OF ED|CHURCH|INC\b|LLC|CORP',re.I)
pw=re.compile(r'PUBLIC ?WORKS|STREET DEP|STREETS|HIGHWAY|HWY|ROAD DIST|ROAD COMM|ENGINEER|TRANSPORT|DPW|\bPW|TRAFFIC|MAINTENANCE',re.I)
safety=re.compile(r'POLICE|PD@|FD@|SHERIFF|\bPD\b|FIRE|EMA\b|EMERGENCY|AMBULANCE|EMS|SHERIFF|911|ESDA|DISPATCH|PSAP|CHIEF|ETSB|RESCUE',re.I)
def norm(n):
    n=n.upper(); n=re.sub(r'[^A-Z ,]',' ',n)
    m=re.match(r'\s*(CITY|VILLAGE|TOWN|TOWNSHIP|COUNTY) OF ([A-Z ]+)',n)
    if m: base,k=m.group(2),m.group(1)
    else:
        m=re.match(r'([A-Z ]+?),\s*(CITY|VILLAGE|TOWN|TOWNSHIP|COUNTY) OF',n) or re.match(r'([A-Z ]+?)\s+(CITY|VILLAGE|TOWN|TOWNSHIP|COUNTY|TWP|CNTY)\b',n)
        if not m: return None
        base,k=m.group(1),m.group(2)
    k={'TWP':'TOWNSHIP','CNTY':'COUNTY'}.get(k,k)
    base=re.sub(r'\b(POLICE|FIRE|DEPT|DEPARTMENT|PUBLIC|WORKS|EMA|HIGHWAY|THE)\b','',base)
    return k.title()+' of '+' '.join(base.split()).title()
rows={}
for l in open(F+'EN.dat',encoding='latin-1'):
    p=l.rstrip('\n').split('|')
    if len(p)<24 or p[17]!='IL' or p[23]!='G' or p[5]!='L': continue
    name=p[7]; em=p[14].strip()
    if not em or '@' not in em or not kind.search(name) or bad.search(name): continue
    h=hd.get(p[1])
    if not h or h[0]!='A': continue
    g=d(h[2]); la=d(h[4]) or g
    if not la or la.year<2021: continue
    key=norm(name)
    if not key: continue
    att=p[20]; addr=p[15]
    txt=att+' '+em+' '+name
    dept='Police / fire / EMA' if safety.search(em+' '+att) else ('Public works / highway' if pw.search(txt) else ('Police / fire / EMA' if safety.search(txt) else 'Town office'))
    score=({'Public works / highway':2,'Town office':1}.get(dept,0), la)
    r=dict(usi=p[1],town=key,dept=dept,name=name.strip(),city=p[16].title(),zip=p[18][:5],phone=p[12],email=em.lower(),contact=att.strip(),address=addr.strip(),callsign=p[4],service=h[1],last_action=la.strftime('%Y-%m-%d'),frn=p[22],pw=score[0]==2,score=score)
    if key not in rows or score>rows[key]['score']: rows[key]=r
out=sorted(rows.values(),key=lambda r:(not r['pw'],r['name']))
for r in out: r.pop('score')
json.dump(out,open(F+'il_towns.json','w'),indent=1)
print(len(out),sum(r['pw'] for r in out))
from collections import Counter
print(Counter(r['dept'] for r in out))
for r in out[:8]: print(r)
