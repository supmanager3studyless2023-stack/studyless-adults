import json,os,time,threading
from concurrent.futures import ThreadPoolExecutor
import pf
det=json.load(open('pf_course_details.json'))
want=[]
seen=set()
for v in det.values():
    if True:
        cp=(v.get('cp') or {}).get('id','')
        if cp.startswith('contact:') and cp not in seen:
            seen.add(cp); want.append(cp.split(':')[1])
print('контактів',len(want),flush=True)
OUT='pf_contacts.json'
data=json.load(open(OUT)) if os.path.exists(OUT) else {}
todo=[c for c in want if c not in data]
lock=threading.Lock(); n=[0]
def work(c):
    if pf.state['remaining'] is not None and pf.state['remaining']<pf.FLOOR: return
    d=pf.req(f"contact/{c}?fields=id,name,midname,lastname,email,phones,birthDate,telegram,description")
    ct=d.get('contact')
    if ct is None: 
        with lock: data[c]={'err':str(d)[:100]}
        return
    with lock:
        data[c]={'name':ct.get('name'),'mid':ct.get('midname'),'last':ct.get('lastname'),'email':ct.get('email'),
                 'phones':[p.get('number') for p in ct.get('phones',[])],'bd':(ct.get('birthDate') or {}).get('date'),'tg':ct.get('telegram')}
        n[0]+=1
        if n[0]%200==0:
            json.dump(data,open(OUT,'w'),ensure_ascii=False); print('done',n[0],pf.state['remaining'],flush=True)
with ThreadPoolExecutor(6) as ex: list(ex.map(work,todo))
json.dump(data,open(OUT,'w'),ensure_ascii=False)
print('DONE',len(data),pf.state['remaining'],flush=True)
