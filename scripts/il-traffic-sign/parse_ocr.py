import re,glob,json,os
D='/tmp/claude-0/-home-user-notion-goapl/3894af6e-7006-5c78-8a81-d4969e11e621/scratchpad/il/'
C=json.load(open(D+'contracts.json'))
LID={'09/18/2026':'65d994df-81db-4009-9161-3614364e03e7','07/31/2026':'b60c6247-ee50-416d-a1b5-38fd7949536f','06/12/2026':'6ccbbaa6-54f8-4eb4-b2eb-1359f5db32b5','04/24/2026':'0a2d2a61-d2f8-4c3e-a90b-aa7f56babf4c'}
known={}
for x in C:
    k=(LID[x['letting']],x['item_contract'].split('-')[-1].upper()); known.setdefault(k,set()).update(b['vendor'] for b in x['bidders'])
res={}
for lid in LID.values():
    cur=None; vend=None
    for f in sorted(glob.glob(D+f'ocr/{lid}/*.ocr.txt')):
        for line in open(f,errors='ignore'):
            m=re.search(r'Contract:?\s*([0-9A-Z]{5})\b',line)
            if m and 'Letting' in line:
                cur=(lid,m.group(1).upper().replace('O','0') if m.group(1)[:2].isdigit() else m.group(1).upper()); res.setdefault(cur,[]); vend=None; continue
            if cur is None: continue
            m=re.match(r'\s*(\d{4})\s',line)
            if m and m.group(1) in known.get(cur,()):
                vend=m.group(1)
                if vend not in [v for v,_ in res[cur]]: res[cur].append([vend,None])
                else: vend=None
            if vend:
                e=[e for e in res[cur] if e[0]==vend][0]
                if e[1] is None:
                    a=re.search(r'\$\s?([\d,]+\.\d{2})',line)
                    if a: e[1]=float(a.group(1).replace(',',''))
                    elif 'No Bid' in line or 'NoBid' in line: e[1]='NB'
json.dump({f'{k[0]}|{k[1]}':v for k,v in res.items()},open(D+'tabs.json','w'),indent=0)
miss=[k for k in known if k not in res]
print('contracts in tabs',len(res),'bidders',sum(len(v) for v in res.values()),'unparsed amounts',sum(1 for v in res.values() for e in v if e[1] is None),'contracts not found',len(miss))
