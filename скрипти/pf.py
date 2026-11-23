import json,subprocess,os
FLOOR=41100  # зупинка, коли x-ratelimit-remaining нижче (бюджет 25 тис. від 66138)
TOK=os.environ['PFT']
H=['-H','Authorization: Bearer '+TOK,'-H','Accept: application/json','-H','Content-Type: application/json']
BASE='https://study-less.planfix.com/rest/'
state={'remaining':None,'calls':0}
def req(path,body=None,m=120):
    if state['remaining'] is not None and state['remaining']<FLOOR:
        raise SystemExit('PlanFix budget floor reached: remaining=%s'%state['remaining'])
    cmd=['curl','-s','-m',str(m),'-D','-']+H+(['-X','POST','-d',json.dumps(body)] if body is not None else [])+[BASE+path]
    out=subprocess.run(cmd,capture_output=True,text=True).stdout
    import re
    parts=re.split(r'\r?\n\r?\n',out,maxsplit=1)
    head=parts[0]; bodytxt=parts[1] if len(parts)>1 else ''
    while bodytxt.startswith('HTTP/'):
        parts=re.split(r'\r?\n\r?\n',bodytxt,maxsplit=1)
        head=parts[0]; bodytxt=parts[1] if len(parts)>1 else ''
    for line in head.splitlines():
        if line.lower().startswith('x-ratelimit-remaining'):
            try: state['remaining']=int(line.split(':')[1])
            except: pass
    state['calls']+=1
    try: return json.loads(bodytxt,strict=False)
    except: return {'raw':bodytxt[:300]}
