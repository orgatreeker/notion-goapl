"""Print the next N clean, VALID, not-yet-sent town emails from the sheet as JSON.
Usage: python3 next_batch.py <xlsx> <verification.csv> <N>"""
import sys,csv,re,json,openpyxl
f,vcsv,n=sys.argv[1],sys.argv[2],int(sys.argv[3])
V={x['Email'].strip().lower():x['validation result'] for x in csv.DictReader(open(vcsv,encoding='utf-8-sig'))}
wb=openpyxl.load_workbook(f)
sent={(r[3] or '').lower() for r in wb['Sent log'].iter_rows(min_row=2,values_only=True)}
T=wb['Towns buying']; h=[c.value for c in T[1]]; x={k:i for i,k in enumerate(h)}
bad=re.compile(r'support|911|kencom|mfd|wastew|communications|police|fire|dispatch|pd\.|pd@|sheriff|ema|esda|hospital|general|school|sped|special|tech|it director|infotech|\bit\b|radio|fcc|bi-county|etsb|rescue|ambulance|ems|marshallcounty|stephensoncounty',re.I)
good=re.compile(r'public works|street|highway|road|engineer|superintendent|administrator|manager|mayor|clerk|trustee|president|village|city',re.I)
c=[]
for r in T.iter_rows(min_row=2,values_only=True):
    e=(r[x['Email']] or '').lower()
    if not e or e in sent or V.get(e)!='VALID' or r[x['Department on licence']]=='Police / fire / EMA' or str(r[x['Email check']]).startswith('Other') or r[0]=='Village of Palatine': continue
    if bad.search(e+' '+str(r[x['Contact on licence']])+' '+str(r[x['Address']])): continue
    s,b=r[x['Email 1']].split('\n\n',1)
    c.append(((0 if good.search(str(r[x['Contact on licence']])+' '+e) else 1,0 if r[x['Email check']]=='Town domain' else 1),dict(name=r[0],to=e,subject=s.replace('Subject: ',''),body=b.strip())))
c.sort(key=lambda z:z[0])
print(json.dumps([z[1] for z in c[:n]],indent=1))
print(f'# remaining clean VALID towns: {len(c)}',file=sys.stderr)
