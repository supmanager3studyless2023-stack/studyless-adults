import json,re,datetime,collections
from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Alignment
exec(open('analyze_adults.py').read().split("m=collections.Counter(); matched={}")[0])
def num(x):
    try: return float(str(x).replace(' ','').replace('\xa0','').replace(',','.'))
    except: return None
def vd(x):
    if not x: return None
    if x.year<2018 or x.year>2030: return None
    return x
persons=dict(by)
emails={};phones={}
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
# --- match table deals
matched={}
for ri,r in rows:
    e=g(r,16).lower(); p=nph(g(r,17))
    c=emails.get(e) or (phones.get(p) if p else None)
    if c: matched.setdefault(c,[]).append((ri,r))
TAR={}
import ast
# --- persons list
def card_start(card): return vd(pd_(card['cf'].get('11635'))) or vd(pd_(card['cf'].get('10307'))) or vd(pd_(card.get('dt')))
plist=[]
for c,cards in persons.items():
    cur=[x for x in cards if x['cf'].get('10433') in sel_status] or list(cards)
    cur.sort(key=lambda x:((x['cf'].get('10433')!='Активний'), -(card_start(x) or datetime.date(2000,1,1)).toordinal()))
    cur=cur[0]
    firstd=[card_start(x) for x in cards if card_start(x)]
    for ri,r in matched.get(c,[]):
        dd_=d(g(r,18)) or d(g(r,10))
        if dd_: firstd.append(dd_)
    plist.append((min(firstd) if firstd else datetime.date(2100,1,1),c,cur))
plist.sort(key=lambda x:(x[0],int(x[1])))
students=[];pay=[]
sid=0
STAT={'Активний':'Активний','Очікує старту':'Очікує старту','На заморозці':'На заморозці'}
for _,c,cur in plist:
    sid+=1
    ct=con.get(c) or {}
    last=(ct.get('last') or '').strip(); nm=(ct.get('name') or '').strip(); mid=(ct.get('mid') or '').strip()
    if not last and nm and ' ' in nm: last,nm=nm.split(' ',1)[0],nm.split(' ',1)[1]
    ph=[p for p in (ct.get('phones') or []) if p]
    em=(ct.get('email') or '').strip() or cur['cf'].get('10413','')
    cf=cur['cf']
    students.append([sid,last,nm,mid,int(c),cur['id'],pd_(ct.get('bd')),cf.get('10499',''),'',', '.join(ph[:2]),em,ct.get('tg') or '',
                     '',(cf.get('10219') or (cur['ass'][0] if cur['ass'] else '')),cf.get('10397',''),cf.get('11645',''),(STAT.get(cf.get('10433')) or cf.get('10433') or cur['status'] or 'Невідомо'),cf.get('10377',''),card_start(cur),vd(pd_(cf.get('10467'))) or vd(pd_(cf.get('11959'))),'; '.join(cur['ass']),'додано 23.11.2026 з PlanFix (картка курсу '+str(cur['id'])+')'])
    full=(' '.join(x for x in (last,nm,mid) if x)).strip()
    items=[]
    for card in persons[c]:
        cf=card['cf']
        pdt=vd(pd_(cf.get('11981'))) or vd(pd_(cf.get('10629'))) or vd(pd_(cf.get('11669'))) or vd(pd_(cf.get('11927')))
        paid=num(cf.get('11381')); tot=num(cf.get('11655')); co=num(cf.get('10791'))
        if paid: s=paid
        elif tot and cf.get('11877') in ('повна',None,''): s=tot
        elif co: s=co
        else: s=tot
        st=card_start(card); fn=vd(pd_(cf.get('10467'))) or vd(pd_(cf.get('11959')))
        note='PlanFix, картка курсу %s; статус активності: %s; статус задачі: %s'%(card['id'],cf.get('10433',''),card['status'])
        if cf.get('11919'): note+='; статус оплати: '+cf['11919']
        if cf.get('11877'): note+='; оплата: '+cf['11877']
        if tot and s!=tot: note+='; загальна вартість пакета %d'%tot
        if co and s!=co: note+='; сума платежу Ч/О %d'%co
        conf='Висока' if (pdt and s and st and fn) else ('Середня' if (s and st) else 'Низька')
        items.append(dict(date=pdt,sum=s,tar=cf.get('10397',''),start=st,fin=fn,src='PlanFix',note=note,conf=conf,exact='так' if pdt else 'ні',ref=card['id']))
    cards_ds=[(x['start'],x['fin'],x['date']) for x in items]
    for ri,r in matched.get(c,[]):
        s0=d(g(r,10)); f0=d(g(r,11)); p0=d(g(r,18))
        dup=False
        for cs,cf_,cp in cards_ds:
            if (s0 and cs and abs((s0-cs).days)<=20) or (f0 and cf_ and abs((f0-cf_).days)<=20) or (p0 and cp and abs((p0-cp).days)<=20): dup=True;break
        if dup: continue
        stg=g(r,25)
        s=num(g(r,21)) or num(g(r,20))
        note='таблиця «Оплати оригінал» (Kommo ID %s); етап угоди: %s'%(g(r,0),stg)
        if g(r,20): note+='; загальна сума %s'%g(r,20)
        if g(r,24): note+='; оплата: '+g(r,24)
        if stg.startswith('Closed - lost (Дубль)') or 'Дубль' in stg:
            s=0; note+='; дубль угоди, суму обнулено'
        if 'Повернення' in stg: note+='; повернення коштів'
        conf='Середня' if (s and (s0 or f0)) else 'Низька'
        items.append(dict(date=p0,sum=s,tar=g(r,12),start=s0,fin=f0,src='Оплати оригінал (Kommo)',note=note,conf=conf,exact='так' if p0 and not (p0.day==1 and p0.month==1) else 'ні',ref=g(r,0)))
    items.sort(key=lambda x:(x['date'] or x['start'] or datetime.date(2100,1,1)))
    lastidx=len(items)-1
    for i,it in enumerate(items):
        pay.append([None,sid,full,it['date'],it['sum'],it['tar'],it['tar'],None,it['start'],it['fin'],'так' if i==lastidx else 'ні',it['exact'],it['conf'],it['src'],it['note']])

# --- table-only persons (немає картки PlanFix)
used={ri for c,dl in matched.items() for ri,r in dl}
import collections as _c
par2={}
def f2(x):
    while par2.setdefault(x,x)!=x:
        par2[x]=par2[par2[x]]; x=par2[x]
    return x
def keys2(r):
    ks=[]
    e=g(r,16).lower().strip()
    if '@' in e: ks.append('e:'+e)
    p=nph(g(r,17))
    if p: ks.append('p:'+p)
    n=re.sub(r'[^а-яіїєґa-z ]','',g(r,1).lower()); n=' '.join(sorted(n.split()))
    if not ks and n: ks.append('n:'+n)
    return ks
rest=[(ri,r) for ri,r in rows if ri not in used]
for ri,r in rest:
    ks=keys2(r) or ['row:%d'%ri]
    for k in ks[1:]: par2[f2(k)]=f2(ks[0])
    f2(ks[0])
grp=_c.defaultdict(list)
for ri,r in rest:
    ks=keys2(r) or ['row:%d'%ri]
    grp[f2(ks[0])].append((ri,r))
glist=[]
for k,dl in grp.items():
    dts=[d(g(r,18)) or d(g(r,10)) for _,r in dl]; dts=[x for x in dts if x]
    glist.append((min(dts) if dts else datetime.date(2100,1,1),k,dl))
glist.sort(key=lambda x:(x[0],x[1]))
ACT=re.compile(r'^(\d+ тиждень|Займається індивідуально|2-6 курс|Допоміжний етап|Фінальний фідбек)$')
for _,k,dl in glist:
    sid+=1
    dl.sort(key=lambda x:(d(g(x[1],18)) or d(g(x[1],10)) or datetime.date(2100,1,1)))
    r0=dl[-1][1]
    flagged=g(r0,37)=='Не активний'; stg=g(r0,25)
    if not flagged and ACT.match(stg): stt='Навчається (лише за таблицею)'
    elif not flagged and stg in ('Треба познайомитись','Невизначена дата старту','Визначена дата старту','Підготовка до старту','Купили офер'): stt='Очікує старту (лише за таблицею)'
    elif 'Closed' in stg or 'Закрыто' in stg: stt='Закрито / відмова (за таблицею)'
    else: stt='Не активний (за таблицею)'
    full=g(r0,1)
    students.append([sid,full,'','','','',d(g(r0,27)),g(r0,28),'',g(r0,17),g(r0,16),g(r0,9),g(r0,3),g(r0,2),g(r0,12),g(r0,31),stt,g(r0,7),d(g(r0,10)),d(g(r0,11)),g(r0,4),'додано 23.11.2026 з таблиці «Оплати оригінал» (картки PlanFix не знайдено); етап угоди: '+stg])
    for i,(ri,r) in enumerate(dl):
        s0=d(g(r,10)); f0=d(g(r,11)); p0=d(g(r,18)); stg2=g(r,25)
        s_=num(g(r,21)) or num(g(r,20))
        note='таблиця «Оплати оригінал» (Kommo ID %s); етап угоди: %s'%(g(r,0),stg2)
        if g(r,20): note+='; загальна сума %s'%g(r,20)
        if g(r,24): note+='; оплата: '+g(r,24)
        if 'Дубль' in stg2: s_=0; note+='; дубль угоди, суму обнулено'
        if 'Повернення' in stg2: note+='; повернення коштів'
        conf='Середня' if (s_ and (s0 or f0)) else 'Низька'
        pay.append([None,sid,full,p0,s_,g(r,12),g(r,12),None,s0,f0,'так' if i==len(dl)-1 else 'ні','так' if p0 and not (p0.day==1 and p0.month==1) else 'ні',conf,'Оплати оригінал (Kommo)',note])

for i,p in enumerate(pay,1): p[0]=i
print('студентів',len(students),'оплат',len(pay))
import pickle
pickle.dump((students,pay),open('adults_data.pkl','wb'))
# --- workbook
wb=Workbook()
hdr=Font(bold=True); fill=PatternFill('solid',fgColor='DDEBF7')
def sheet(name,header,widths=None):
    ws=wb.create_sheet(name); ws.append(header)
    for c in ws[1]: c.font=hdr; c.fill=fill; c.alignment=Alignment(wrap_text=True,vertical='top')
    ws.freeze_panes='C2' if name in ('Студенти','Оплати','Аналітика') else 'A2'
    return ws
wb.remove(wb.active)
rd=sheet('Читай мене',['Дорослий · Зведення студентів і оплат'])
txt=["Оновлено 23.11.2026. Зібрано з PlanFix (картки курсу «Adults. Картка курсу») і з таблиці «студенти новий формат 10.10.22» (вкладка «Оплати оригінал»). У джерелах нічого не змінювалось.","",
"ЩО ЦЕ","Зведення дорослих студентів за зразком Kids: «Студенти», «Оплати» (рядок = оплата/курс), «Аналітика» (формули).",
"Завантажено всіх студентів: %d осіб, %d рядків оплат. З PlanFix: усі студенти з картками курсу (статус із поля «Статус активності»). З таблиці «Оплати оригінал»: студенти без картки PlanFix (статус «лише за таблицею»)."%(len(students),len(pay)),"",
"ВКЛАДКИ","Студенти: ПІБ, контакти, ДН, країна, поточна картка курсу, статус навчання, тариф, наставник.",
"Оплати: колонки як у Kids. Джерело «PlanFix» це картка курсу (дата оплати курсу, сума, початок, завершення). «Оплати оригінал (Kommo)» це історія 2022–2025 із таблиці, зіставлена зі студентом за поштою або телефоном.",
"Аналітика: формули по «Оплати» (LT, пауза, перекриття).","Кейси для менеджерів і Непривʼязані оплати: порожні, наповнюються під час звірки.","",
"ЯК ЧИТАТИ","Сума в PlanFix: «Загальна оплачена сума», інакше «Загальна вартість пакета» (якщо «повна»), інакше «Сума платежу Ч/О». Це не завжди факт оплати, дивись примітку й «Впевненість».",
"Дати старту/фінішу в PlanFix, що менші за 2018 або більші за 2030 (заглушки 2050/2100), не використовуються.",
"Статус «Активний/Очікує старту» береться з поля «Статус активності» картки курсу, а не зі статусу задачі.","",
"ЩО ПОКИ НЕ ПОВНЕ","Оплати PlanFix для 2026 року з задач «Облік оплат», STRIPE, WFP ще не додані.","Таблиця-джерело ведеться до 2025 року, тому історія 2026 є лише в PlanFix.","Студенти без вказаного наставника в PlanFix (≈60% активних) мають порожню колонку «Наставник».","Кейси для менеджерів і звірка з Kommo не проводились."]
for t in txt: rd.append([t])
rd.column_dimensions['A'].width=140
for r in rd.iter_rows(min_row=2):
    r[0].alignment=Alignment(wrap_text=True)
rd['A1'].font=Font(bold=True,size=14)
ws=sheet('Студенти',['ID студента','Прізвище','Імʼя','По батькові','ID контакту PlanFix','ID картки курсу (поточної)','Дата народження','Країна','Місто','Телефон','Пошта','Телеграм','Sales менеджер','Наставник','Поточний тариф','Тип тарифу','Статус навчання','Курс','Дата старту (поточна)','Дата фінішу (поточна)','Виконавці картки PlanFix','Примітка'])
for s in students:
    ws.append(s)
for row in ws.iter_rows(min_row=2):
    for i in (6,18,19): row[i].number_format='dd.mm.yyyy'
wo=sheet('Оплати',['ID оплати','ID студента','ПІБ (для перевірки)','Дата оплати','Сума оплати','Тариф (як у джерелі)','Тариф (актуальна назва)','Кількість (тижнів/уроків/міс.)','Початок оплаченого періоду','Завершення оплаченого періоду','Остання оплата?','Дата точна?','Впевненість','Джерело','Примітка','Позначка періоду','Початок (без перекриття)','Кінець (без перекриття)'])
N=len(pay)+1
for i,p in enumerate(pay):
    r=i+2
    row=p+[
     f'=IF(J{r}<>"";"";IF(OR(AND(E{r}<>"";E{r}>0;E{r}<=1500);F{r}="Бронь";F{r}="Інше");"бронь / застава";"період невідомий"))',
     f'=IF(OR(I{r}="";J{r}="");I{r};IF(R{r}=J{r};I{r};R{r}-IFERROR(IF(VALUE(SUBSTITUTE(H{r};",";"."))>24;J{r}-I{r}+1;ROUND(VALUE(SUBSTITUTE(H{r};",";"."))*28));J{r}-I{r}+1)+1))',
     f'=IF(OR(I{r}="";J{r}="");J{r};IF(AND(B{r+1}=B{r};Q{r+1}<>"");IF(J{r}>=Q{r+1};Q{r+1}-1;J{r});J{r}))']
    wo.append(row)
for row in wo.iter_rows(min_row=2):
    for i in (3,8,9,16,17): row[i].number_format='dd.mm.yyyy'
    row[4].number_format='#,##0'
wa=sheet('Аналітика',['ID студента','ПІБ','Статус','Дата першої покупки','Дата останньої покупки','Кількість покупок','Сума покупок, грн','Початок першого оплаченого періоду','Кінець останнього оплаченого періоду','Календарний LT, тижнів','Оплачений LT, тижнів','Пауза, тижнів','Поточний тариф','Курс','Повторний покупець?','Примітка'])
for i in range(len(students)):
    r=i+2
    wa.append([f"='Студенти'!A{r}",f"=TRIM('Студенти'!B{r}&\" \"&'Студенти'!C{r}&\" \"&'Студенти'!D{r})",f"='Студенти'!Q{r}",
     f"=IF(F{r}=0;\"\";MINIFS('Оплати'!$D:$D;'Оплати'!$B:$B;A{r};'Оплати'!$E:$E;\">0\"))",
     f"=IF(F{r}=0;\"\";MAXIFS('Оплати'!$D:$D;'Оплати'!$B:$B;A{r};'Оплати'!$E:$E;\">0\"))",
     f"=COUNTIFS('Оплати'!$B:$B;A{r};'Оплати'!$E:$E;\">0\")",
     f"=SUMIFS('Оплати'!$E:$E;'Оплати'!$B:$B;A{r})",
     f"=IF(F{r}=0;\"\";MINIFS('Оплати'!$I:$I;'Оплати'!$B:$B;A{r};'Оплати'!$E:$E;\">0\"))",
     f"=IF(F{r}=0;\"\";IF(MAXIFS('Оплати'!$J:$J;'Оплати'!$B:$B;A{r};'Оплати'!$E:$E;\">0\")=0;\"\";MAXIFS('Оплати'!$J:$J;'Оплати'!$B:$B;A{r};'Оплати'!$E:$E;\">0\")))",
     f"=IF(OR(H{r}=\"\";I{r}=\"\");\"\";IF(I{r}<=H{r};\"\";ROUND((I{r}-H{r}+1)/7;1)))",
     f"=IF(F{r}=0;\"\";IF(J{r}=\"\";\"\";ROUND(SUMPRODUCT(('Оплати'!$B$2:$B${N}=A{r})*('Оплати'!$E$2:$E${N}>0)*('Оплати'!$Q$2:$Q${N}>0)*('Оплати'!$R$2:$R${N}>'Оплати'!$Q$2:$Q${N})*('Оплати'!$R$2:$R${N}-'Оплати'!$Q$2:$Q${N}+1))/7;1)))",
     f"=IF(OR(J{r}=\"\";K{r}=\"\");\"\";MAX(0;ROUND(J{r}-K{r};1)))",
     f"='Студенти'!O{r}",f"='Студенти'!R{r}",f"=IF(F{r}>1;\"так\";IF(F{r}=1;\"ні\";\"\"))",
     f"=IF(F{r}=0;\"без платних оплат\";IF(J{r}=\"\";\"немає завершеного періоду, LT не рахується\";IF(COUNTIFS('Оплати'!$B:$B;A{r};'Оплати'!$P:$P;\"період невідомий\")>0;\"є оплата без кінця періоду, LT може бути занижений\";IF(K{r}>J{r}+1;\"періоди перекриваються, оплачений LT завищений\";\"\"))))"])
for row in wa.iter_rows(min_row=2):
    for i in (3,4,7,8): row[i].number_format='dd.mm.yyyy'
    row[6].number_format='#,##0'
sheet('Кейси для менеджерів',['№','ID студента','ПІБ','Поточний менеджер','Чого не вистачає','Що вже перевірено і знайдено','Питання для того, хто розбирає','Відповідь (заповнює менеджер)','Статус','Кого залучити','Посилання PlanFix'])
sheet('Непривʼязані оплати',['Дата оплати','Сума, грн','Ім\'я в джерелі','Тариф','Кількість','Джерело і рядок','Тип випадку','Можливий студент (ID · ім\'я)','Відповідь (ID студента, «новий» чи «ігнор»)','Статус','Примітка'])
widths={'Студенти':[9,18,14,16,12,14,12,14,8,18,28,14,14,20,22,12,14,10,12,12,24,40],'Оплати':[9,9,28,12,11,20,20,10,12,12,9,9,11,22,60,16,12,12]}
for n,ws_ in widths.items():
    for i,w in enumerate(ws_): wb[n].column_dimensions[chr(65+i)].width=w
wb.save('Дорослий_зведення_оплат_повний.xlsx')
print('saved')
