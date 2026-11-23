import json,re,datetime,collections
SRC='kommo_oplaty_original.json'
V=json.load(open(SRC))['values']
rows=[(i+3,r) for i,r in enumerate(V[2:]) if len(r)>1 and r[0].strip()]
g=lambda r,i: r[i].strip() if len(r)>i else ''
def d(x):
    x=x.strip()
    for f in ('%d.%m.%Y','%d.%m.%y'):
        try:
            dt=datetime.datetime.strptime(x,f).date()
            if dt.year in (1970,2050) or dt.year<2018 or dt.year>2027: return None
            return dt
        except: pass
    return None
def num(x):
    try: return float(x.replace(' ','').replace('\xa0','').replace(',','.'))
    except: return None
ACTIVE=re.compile(r'^(\d+ тиждень|Займається індивідуально|2-6 курс|Допоміжний етап|Фінальний фідбек)$')
WAIT={'Треба познайомитись','Невизначена дата старту','Визначена дата старту','Підготовка до старту','Купили офер'}
def cls(r):
    if g(r,37)=='Не активний': return None
    st=g(r,25)
    if ACTIVE.match(st): return 'Навчається'
    if st in WAIT: return 'Очікує старту'
    return None
# union-find on email/phone
par={}
def find(x):
    while par.setdefault(x,x)!=x:
        par[x]=par[par[x]]; x=par[x]
    return x
def union(a,b): par[find(a)]=find(b)
def keys(r):
    ks=[]
    e=g(r,16).lower().strip()
    if '@' in e: ks.append('e:'+e)
    p=re.sub(r'\D','',g(r,17))
    if len(p)>=9: ks.append('p:'+p[-9:])
    n=re.sub(r'[^а-яіїєґa-z ]','',g(r,1).lower()); n=' '.join(sorted(n.split()))
    if not ks and n: ks.append('n:'+n)
    return ks
for ri,r in rows:
    ks=keys(r) or ['row:'+str(ri)]
    for k in ks[1:]: union(ks[0],k)
    find(ks[0])
persons=collections.defaultdict(list)
for ri,r in rows:
    ks=keys(r) or ['row:'+str(ri)]
    persons[find(ks[0])].append((ri,r))
print('угод',len(rows),'осіб',len(persons))
sel=[]
cnt=collections.Counter()
for pk,deals in persons.items():
    cs=[cls(r) for _,r in deals]
    if any(c=='Навчається' for c in cs): s='Навчається'
    elif any(c=='Очікує старту' for c in cs): s='Очікує старту'
    else: continue
    sel.append((pk,s,deals)); cnt[s]+=1
print(cnt,'угод у вибраних',sum(len(x[2]) for x in sel))
json.dump({'n':len(sel)},open('sel_meta.json','w'))
import pickle
pickle.dump(sel,open('sel.pkl','wb'))
