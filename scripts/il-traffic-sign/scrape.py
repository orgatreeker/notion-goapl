import re,json,html,subprocess,sys,os,concurrent.futures as cf
D='/tmp/claude-0/-home-user-notion-goapl/3894af6e-7006-5c78-8a81-d4969e11e621/scratchpad/il/'
BASE='https://webapps1.dot.illinois.gov'
LET={'09/18/2026':'65d994df-81db-4009-9161-3614364e03e7','07/31/2026':'b60c6247-ee50-416d-a1b5-38fd7949536f',
     '06/12/2026':'6ccbbaa6-54f8-4eb4-b2eb-1359f5db32b5','04/24/2026':'0a2d2a61-d2f8-4c3e-a90b-aa7f56babf4c'}
def get(url,fn):
    p=D+'pages/'+fn
    if not os.path.exists(p) or os.path.getsize(p)<1000:
        subprocess.run(['curl','-sSL','--retry','3','-o',p,url],check=False)
    return open(p,encoding='utf-8',errors='ignore').read()
os.makedirs(D+'pages',exist_ok=True)
def clean(x): return re.sub(r'\s+',' ',html.unescape(re.sub(r'<[^>]+>',' ',x))).strip()
def lis(s,header):
    i=s.find('<th>'+header+'</th>')
    if i<0: return []
    # find nth column: count th before
    thead=s[s.rfind('<thead',0,i):s.find('</thead>',i)]
    cols=re.findall(r'<th>(.*?)</th>',thead); idx=cols.index(header)
    body=s[s.find('<tbody>',i):s.find('</tbody>',i)]
    tds=re.findall(r'<td>(.*?)</td>',body,re.S)
    if idx>=len(tds): return []
    v=re.findall(r'<li>(.*?)</li>',tds[idx],re.S)
    return [clean(x) for x in v] if v else [clean(tds[idx])]
def parse(s):
    r={}
    m=re.search(r'Letting\s+\d\d:\d\d [AP]M\s*</[^>]+>\s*(?:<[^>]+>\s*)*([0-9A-Z]{3}-[0-9A-Z]{5})',s)
    t=clean(s)
    m=re.search(r'Letting \d\d:\d\d [AP]M (\S+-\S+)',t); r['item_contract']=m.group(1) if m else ''
    m=re.search(r'Bulletin Description (.*?) Region District',t); r['desc']=m.group(1) if m else ''
    for h,k in [('County(s)','county'),('Marked Route(s)','route'),('Section(s)','section'),('District','district')]:
        r[k]=', '.join(lis(s,h))
    j=s.find('Authorized Bidders</h3>'); seg=s[j:s.find('</ul>',j)] if j>0 else ''
    bs=[]
    for li in re.findall(r'<li>(.*?)</li>',seg,re.S):
        name=clean((re.search(r'<strong>(.*?)</strong>',li,re.S) or [None,''])[1])
        num,_,name=name.partition(' ')
        em=re.search(r'Email:\s*([^<\s]+)',li); ph=re.search(r'Phone:\s*(.*?)\s*Fax:',li,re.S)
        amt=re.search(r'alert-success">\s*\$([\d,\.]+)',li)
        addr=li.split('</strong>')[1].split('Phone:')[0] if '</strong>' in li else ''
        addr=', '.join(x for x in (clean(a) for a in re.split(r'<br\s*/?>',addr)) if x)
        bs.append(dict(vendor=num,company=name,email=(em.group(1) if em and '@' in em.group(1) else ''),phone=clean(ph.group(1)) if ph else '',address=addr,bid=float(amt.group(1).replace(',','')) if amt else None))
    r['bidders']=bs
    return r
out=[]
for date,lid in LET.items():
    s=get(f'{BASE}/WCTB/LbLettingDetail/Index/{lid}',f'let_{lid}.html')
    cids=sorted(set(re.findall(r'contractId=([0-9a-f-]{36})',s)))
    print(date,len(cids),file=sys.stderr)
    def work(c):
        p=get(f'{BASE}/WCTB/LbContractDetail/Index/{lid}?page=1&contractId={c}',f'c_{c}.html')
        r=parse(p); r['letting']=date; r['url']=f'{BASE}/WCTB/LbContractDetail/Index/{lid}?page=1&contractId={c}'; return r
    with cf.ThreadPoolExecutor(8) as ex: out+=list(ex.map(work,cids))
json.dump(out,open(D+'contracts.json','w'),indent=1)
print(len(out),sum(len(x['bidders']) for x in out))
