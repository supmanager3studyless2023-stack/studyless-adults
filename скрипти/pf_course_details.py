import json,os,time,pf
IDS=[10433,10507,10307,11635,10467,11959,11355,11449,11931,10915,11455,10397,11305,11875,11645,10609,10753,10377,11343,10219,11297,11093,11515,10817,11655,11373,11381,10791,11399,10629,11981,11669,11927,11877,11919,12021,10311,11979,11965,10303,10853,10413,10499,10501,10837,10839,10321,10445,10111,10511,11361,11363,12005,10505,11379,11581,11245,10551,10541,11173,11179,11181,10099]
fields="id,name,status,dateTime,parent,counterparty,assignees,"+",".join(map(str,IDS))
OUT='pf_course_details.json'
data=json.load(open(OUT)) if os.path.exists(OUT) else {}
off=len(data)
def slim(t):
    r={'id':t['id'],'name':t.get('name'),'status':(t.get('status') or {}).get('name'),'dt':(t.get('dateTime') or {}).get('date'),
       'parent':(t.get('parent') or {}).get('id'),'cp':t.get('counterparty'),
       'ass':[u['name'] for u in (t.get('assignees') or {}).get('users',[])]}
    cf={}
    for c in t.get('customFieldData') or []:
        v=c.get('stringValue')
        if v in (None,''): 
            vv=c.get('value')
            if isinstance(vv,dict): v=vv.get('date') or vv.get('name') or json.dumps(vv,ensure_ascii=False)
            elif vv not in (None,''): v=str(vv)
            else: v=''
        if v!='': cf[str(c['field']['id'])]=v
    r['cf']=cf
    return r
fails=0
while True:
    d=pf.req("task/list",{"offset":off,"pageSize":100,"filters":[{"type":51,"operator":"equal","value":5408463}],"fields":fields},180)
    t=d.get('tasks')
    if t is None:
        fails+=1; print('err',off,str(d)[:200],flush=True)
        if fails>4: break
        time.sleep(5); continue
    fails=0
    for x in t: data[str(x['id'])]=slim(x)
    json.dump(data,open(OUT,'w'),ensure_ascii=False)
    print('off',off,len(data),pf.state['remaining'],flush=True)
    if len(t)<100: break
    off+=100
print('DONE',len(data),pf.state['remaining'],flush=True)
