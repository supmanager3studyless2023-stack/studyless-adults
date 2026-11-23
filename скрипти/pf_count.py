import json,subprocess,os,time,collections
TOK=os.environ['PFT']
H=['-H','Authorization: Bearer '+TOK,'-H','Accept: application/json','-H','Content-Type: application/json']
def req(u,body,m=120):
    r=subprocess.run(['curl','-s','-m',str(m)]+H+['-X','POST','-d',json.dumps(body),u],capture_output=True,text=True).stdout
    try: return json.loads(r)
    except: return {'raw':r[:200]}
base="https://study-less.planfix.com/rest/task/list"
for tpl in (5401417,5408463):
    ids=[];off=0;t0=time.time()
    while True:
        d=req(base,{"offset":off,"pageSize":100,"filters":[{"type":51,"operator":"equal","value":tpl}],"fields":"id,name,status,dateTime"})
        t=d.get('tasks')
        if t is None: print('err',tpl,off,d,flush=True); break
        ids+=[(x['id'],x['name'],(x.get('status') or {}).get('name'),(x.get('dateTime') or {}).get('date')) for x in t]
        if len(t)<100: break
        off+=100
        if off%1000==0: print(tpl,off,round(time.time()-t0),flush=True)
    json.dump(ids,open(f'pf_tasks_{tpl}.json','w'),ensure_ascii=False)
    c=collections.Counter(x[2] for x in ids)
    print('template',tpl,len(ids),c.most_common(40),flush=True)
