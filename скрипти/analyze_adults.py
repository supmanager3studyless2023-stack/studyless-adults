import json,re,datetime,collections,pickle
det=json.load(open('pf_course_details.json')); con=json.load(open('pf_contacts.json'))
def pd_(x):
    if not x: return None
    for f in ('%d-%m-%Y','%d.%m.%Y','%Y-%m-%d'):
        try: return datetime.datetime.strptime(x.strip(),f).date()
        except: pass
    return None
sel_status=('Активний','Очікує старту','На заморозці')
by=collections.defaultdict(list)
for v in det.values():
    cid=(v.get('cp') or {}).get('id','')
    if cid.startswith('contact:'): by[cid.split(':')[1]].append(v)
persons={c:cards for c,cards in by.items() if any(x['cf'].get('10433') in sel_status for x in cards)}
print('осіб',len(persons),'карток курсу в них',sum(len(v) for v in persons.values()))
# google table
exec(open('build_adults.py').read().split("# union-find")[0])
def nph(x): 
    p=re.sub(r'\D','',x); return p[-9:] if len(p)>=9 else None
emails={}; phones={}
for c in persons:
    ct=con.get(c) or {}
    e=(ct.get('email') or '').lower().strip()
    if e: emails.setdefault(e,c)
    for card in persons[c]:
        e2=card['cf'].get('10413','').lower().strip()
        if e2: emails.setdefault(e2,c)
    for p in ct.get('phones') or []:
        k=nph(p or '')
        if k: phones.setdefault(k,c)
m=collections.Counter(); matched={}
for ri,r in rows:
    e=g(r,16).lower(); p=nph(g(r,17))
    c=emails.get(e) or (phones.get(p) if p else None)
    if c: matched.setdefault(c,[]).append((ri,r)); m['угод зіставлено']+=1
    else: m['не зіставлено']+=1
print(m,'осіб із угодами таблиці',len(matched))
# overlap of deals with PF cards
ov=collections.Counter()
for c,deals in matched.items():
    cards=persons[c]
    for ri,r in deals:
        s=d(g(r,10)); pdt=d(g(r,18)); f=d(g(r,11))
        hit=False
        for card in cards:
            cs=pd_(card['cf'].get('11635')); cf=pd_(card['cf'].get('10467')); cp=pd_(card['cf'].get('11981'))
            if (s and cs and abs((s-cs).days)<=20) or (f and cf and abs((f-cf).days)<=20) or (pdt and cp and abs((pdt-cp).days)<=20): hit=True;break
        ov['співпало з карткою PF' if hit else 'тільки в таблиці']+=1
print(ov)
yrs=collections.Counter()
for c,deals in matched.items():
    for ri,r in deals:
        pdt=d(g(r,18)) or d(g(r,10))
        yrs[pdt.year if pdt else None]+=1
print(sorted(yrs.items(),key=lambda x:str(x[0])))
# PF cards by year of start
y2=collections.Counter()
for c,cards in persons.items():
    for card in cards:
        s=pd_(card['cf'].get('11635')) or pd_(card.get('dt'))
        y2[s.year if s else None]+=1
print('PF картки за роком старту',sorted(y2.items(),key=lambda x:str(x[0])))
