import json,os,time,collections
import pf
ids=json.load(open('pf_tasks_5408463.json')) if os.path.exists('pf_tasks_5408463.json') else []
off=len(ids)
fails=0
while True:
    d=pf.req("task/list",{"offset":off,"pageSize":100,"filters":[{"type":51,"operator":"equal","value":5408463}],"fields":"id,name,status"})
    t=d.get('tasks')
    if t is None:
        fails+=1; print('err',off,d,flush=True)
        if fails>5: break
        time.sleep(5); continue
    fails=0
    ids+=[(x['id'],x['name'],(x.get('status') or {}).get('name'),None) for x in t]
    json.dump(ids,open('pf_tasks_5408463.json','w'),ensure_ascii=False)
    if len(t)<100: break
    off+=100
    if off%1000==0: print('off',off,'remaining',pf.state['remaining'],flush=True)
c=collections.Counter(x[2] for x in ids)
print('DONE',len(ids),pf.state['remaining'],c.most_common(40),flush=True)
