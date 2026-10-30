import subprocess,json,re,urllib.parse,concurrent.futures as cf,sys
S={'San Jose':'sanjose','Fresno':'fresno','Sacramento':'sacramento','Oakland':'oakland','Stockton':'stockton','Riverside':'riversideca','San Bernardino':'sanbernardino','Fontana':'fontana','Huntington Beach':'huntingtonbeach','Salinas':'salinas','Hayward':'hayward','Sunnyvale':'sunnyvaleca','Pomona':'pomona','Fullerton':'fullerton','Orange':'cityoforange','Visalia':'visalia','Santa Clara':'santaclara','Temecula':'temeculaca','Murrieta':'murrieta','Costa Mesa':'costamesa','Rialto':'rialto','Hesperia':'hesperia','San Marcos':'sanmarcos','Merced':'cityofmerced','Napa':'napa','Chino':'chino','Mountain View':'mountainview','Alameda':'alameda','Newport Beach':'newportbeach','San Leandro':'sanleandro','Lodi':'lodica','Santa Barbara':'santabarbara','Dublin':'dublin',
'Fresno County':'fresnocounty','San Mateo County':'sanmateocounty','Monterey County':'monterey','Solano County':'solano','San Diego County':'sdcounty','Los Angeles County':'lacounty','Humboldt County':'humboldt','Mendocino County':'mendocino'}
KW=re.compile(r'traffic signal|signali[sz]ation|signal (moderni|upgrade|improve|install|replace|timing|synchron)|pavement marking|striping|re-?striping|thermoplastic|crosswalk|rectangular rapid|RRFB|HAWK|pedestrian hybrid|flashing beacon|sign replacement|street name sign|traffic sign|signage|retroreflect|road diet|safe routes to school|HSIP|highway safety improvement|SS4A|safe streets (and|&) roads|vision zero|traffic calming|speed (hump|table|feedback)|guardrail|intersection improvement',re.I)
since=sys.argv[1] if len(sys.argv)>1 else '2026-08-31'
def get(slug):
    out=[];skip=0
    while True:
        q=f"$filter=MatterIntroDate ge datetime'{since}'&$top=1000&$skip={skip}"
        url=f"https://webapi.legistar.com/v1/{slug}/matters?"+urllib.parse.quote(q,safe='=&$\'')
        try: d=json.loads(subprocess.run(['curl','-s','-m','40',url],capture_output=True,text=True).stdout)
        except Exception: break
        if not isinstance(d,list) or not d: break
        out+=d; skip+=1000
        if len(d)<1000: break
    return out
res=[]
with cf.ThreadPoolExecutor(10) as ex:
    for (name,slug),ms in zip(S.items(),ex.map(get,S.values())):
        hits=[m for m in ms if KW.search((m.get('MatterTitle') or '')+' '+(m.get('MatterName') or ''))]
        print(name,len(ms),len(hits),file=sys.stderr)
        for m in hits: res.append(dict(buyer=name,slug=slug,id=m['MatterId'],file=m.get('MatterFile'),date=m.get('MatterIntroDate'),status=m.get('MatterStatusName'),type=m.get('MatterTypeName'),title=m.get('MatterTitle') or m.get('MatterName'),body=m.get('MatterBodyName')))
json.dump(res,open('matters.json','w'),indent=1)
print(len(res))
