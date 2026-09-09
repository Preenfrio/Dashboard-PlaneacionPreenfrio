#!/usr/bin/env python3
from __future__ import annotations
import io,json,threading,webbrowser
from datetime import datetime,timedelta,time
from http.server import ThreadingHTTPServer,BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse
import pandas as pd

BASE=Path(__file__).resolve().parent; EXCEL=BASE/'PLANEACION_SEMANAL 1.xlsx'; HTML=BASE/'dashboard_preenfrio_v2.html'; HOST,PORT='127.0.0.1',8010
CAP={'NELLY':12.0,'FORTALEZA':10.0}; START_HOUR=19; GAP_MIN=30
STATE={'base':None,'overrides':{},'payload':None,'mtime':None}
def norm(v):return ' '.join(str(v if v is not None else '').strip().upper().split())
def col(df,n,req=True):
 for c in df.columns:
  if norm(c)==norm(n):return c
 if req:raise ValueError(f'Falta columna: {n}')
def asdt(v):
 d=pd.to_datetime(v,dayfirst=True,errors='coerce');return None if pd.isna(d) else d.to_pydatetime()
def hrs(v):
 if pd.isna(v) or v=='':return 0.0
 if isinstance(v,(datetime,time)):return v.hour+v.minute/60+v.second/3600
 if isinstance(v,timedelta):return v.total_seconds()/3600
 if isinstance(v,(int,float)):return float(v)*24 if 0<=float(v)<1.5 else float(v)
 p=str(v).strip().split(':')
 try:return float(p[0])+float(p[1])/60+(float(p[2])/3600 if len(p)>2 else 0) if len(p)>1 else float(p[0])
 except:return 0.0
def catalogue():
 df=pd.read_excel(EXCEL,sheet_name='CATALOGO',engine='openpyxl'); c={n:col(df,n) for n in ['CLIENTE','DESTINO','TRÁNSITO (HRS)','CICLO HORAS FORTALEZA','CICLO HORAS NELLY','CRITICIDAD']};out={}
 for _,r in df.iterrows():
  d=norm(r[c['DESTINO']]);
  if d:out[d]={'client':norm(r[c['CLIENTE']]),'transit_h':hrs(r[c['TRÁNSITO (HRS)']]),'critical':int(float(r[c['CRITICIDAD']])) if not pd.isna(r[c['CRITICIDAD']]) else 99,'cycle':{'NELLY':hrs(r[c['CICLO HORAS NELLY']]),'FORTALEZA':hrs(r[c['CICLO HORAS FORTALEZA']])}}
 return out
def read_jobs():
 cat=catalogue();df=pd.read_excel(EXCEL,sheet_name='PLANEACION',engine='openpyxl');c={n:col(df,n) for n in ['Finca','Destino','Pallets','Fecha de proceso','FECHA CITA']};c['SEMANA']=col(df,'SEMANA',False);c['HORAS TRANSITO']=col(df,'HORAS TRANSITO',False);c['NIVEL CRITICO']=col(df,'NIVEL CRITICO',False)
 out={}
 for i,r in df.iterrows():
  finca=str(r[c['Finca']]).strip();dest=str(r[c['Destino']]).strip();info=cat.get(norm(dest));pdte=asdt(r[c['Fecha de proceso']]);appt=asdt(r[c['FECHA CITA']])
  try:pal=float(r[c['Pallets']])
  except:pal=0
  if not info or not pdte or not appt or pal<=0:continue
  transit=hrs(r[c['HORAS TRANSITO']]) if c['HORAS TRANSITO'] else info['transit_h'];critical=int(float(r[c['NIVEL CRITICO']])) if c['NIVEL CRITICO'] and not pd.isna(r[c['NIVEL CRITICO']]) else info['critical'];week=str(int(float(r[c['SEMANA']]))) if c['SEMANA'] and not pd.isna(r[c['SEMANA']]) else str(pdte.isocalendar().week)
  key=f'{norm(finca)}|{norm(dest)}|{pdte.date()}|{appt.date()}|{info["client"]}'
  if key not in out:out[key]={'key':key,'source':i+2,'farm':finca,'destination':dest,'client':info['client'],'process_date':pdte.replace(hour=0,minute=0,second=0,microsecond=0),'appointment':appt,'transit_h':transit,'critical':critical,'pallets':pal,'week':week,'cycle_options':info['cycle']}
  else:out[key]['pallets']+=pal
 return list(out.values())
def allowed(j):return ['NELLY'] if j['client']=='SMART' else [o for o in ['NELLY','FORTALEZA'] if j['cycle_options'][o]>0]
def pack(origin,seq):
 if not seq:return [],[]
 cap=CAP[origin];cur=datetime.combine(seq[0]['process_date'].date(),time(START_HOUR));rem={j['key']:j['pallets'] for j in seq};first={};last={};cycles=[];i=0
 while i<len(seq):
  space=cap;loads=[];dur=0
  while space>1e-9 and i<len(seq):
   j=seq[i];q=min(space,rem[j['key']]);loads.append({'key':j['key'],'destination':j['destination'],'pallets':q});dur=max(dur,j['cycle_options'][origin]);first.setdefault(j['key'],cur);rem[j['key']]-=q;space-=q
   if rem[j['key']]<=1e-9:i+=1
  end=cur+timedelta(hours=dur);cycles.append({'cycle_no':len(cycles)+1,'origin':origin,'process_date':seq[0]['process_date'],'start':cur,'end':end,'duration_h':dur,'capacity':cap,'used_pallets':cap-space,'free_pallets':space,'loads':loads})
  for x in loads:
   if rem[x['key']]<=1e-9:last[x['key']]=end
  cur=end+timedelta(minutes=GAP_MIN)
 rows=[]
 for order,j in enumerate(seq,1):
  end=last[j['key']];arrival=end+timedelta(hours=j['transit_h']);slack=(j['appointment']-arrival).total_seconds()/3600
  rows.append({**j,'assigned_origin':origin,'order':order,'start':first[j['key']],'end':end,'arrival':arrival,'slack_h':slack,'status':'CUMPLE' if slack>=0 else 'NO CUMPLE'})
 return rows,cycles
def automatic(jobs):
 groups={}
 for j in jobs:groups.setdefault(j['process_date'].date(),[]).append(j)
 assign={}
 for date,grp in groups.items():
  q={'NELLY':[],'FORTALEZA':[]}
  for j in sorted(grp,key=lambda x:(x['appointment'],x['critical'],x['source'])):
   cand=allowed(j);best=None
   for o in cand:
    test=q[o]+[j];rows,_=pack(o,test);score=(sum(x['status']=='NO CUMPLE' for x in rows),sum(max(0,-x['slack_h']) for x in rows),len(test))
    if best is None or score<best[0]:best=(score,o)
   if best:q[best[1]].append(j)
  for o in q:
   for n,j in enumerate(q[o],1):assign[j['key']]={'origin':o,'order':n}
 return assign
def recalc(overrides=None):
 jobs=read_jobs();auto=automatic(jobs);ov=overrides if overrides is not None else STATE['overrides'];groups={}
 for j in jobs:
  if j['key'] not in auto: continue
  a=ov.get(j['key'],auto.get(j['key'],{}));origin=a.get('origin') or auto[j['key']]['origin'];order=int(a.get('order') or auto[j['key']]['order'])
  if origin!='DIRECTO SIN PREENFRIO':
   if j['client']=='SMART':origin='NELLY'
   if origin not in allowed(j):origin=auto[j['key']]['origin']
  groups.setdefault((j['process_date'].date(),origin),[]).append((order,j))
 rows=[];cycles=[]
 for (date,origin),items in sorted(groups.items()):
  seq=[j for _,j in sorted(items,key=lambda x:(x[0],x[1]['appointment'],x[1]['critical'],x[1]['source']))]
  if origin=='DIRECTO SIN PREENFRIO':
   for order,j in enumerate(seq,1):
    dispatch=datetime.combine(j['process_date'].date(),time(START_HOUR));arrival=dispatch+timedelta(hours=j['transit_h']);slack=(j['appointment']-arrival).total_seconds()/3600
    rows.append({**j,'assigned_origin':origin,'order':order,'start':dispatch,'end':dispatch,'arrival':arrival,'slack_h':slack,'status':'CUMPLE' if slack>=0 else 'NO CUMPLE'})
  else:
   r,c=pack(origin,seq);rows+=r;cycles+=c
 return rows,cycles
def serial(rows,cycles):
 def iso(v):return v.isoformat(timespec='minutes') if isinstance(v,datetime) else v
 rf=['key','farm','destination','client','process_date','appointment','transit_h','critical','pallets','week','assigned_origin','order','start','end','arrival','slack_h','status'];cf=['cycle_no','origin','process_date','start','end','duration_h','capacity','used_pallets','free_pallets','loads']
 return {'ok':True,'version':'2.1','excel':EXCEL.name,'modified':datetime.fromtimestamp(EXCEL.stat().st_mtime).isoformat(timespec='seconds'),'data':[{f:iso(x[f]) for f in rf} for x in rows],'cycles':[{f:iso(x[f]) for f in cf} for x in cycles]}
def payload():
 m=EXCEL.stat().st_mtime
 if STATE['payload'] is None or STATE['mtime']!=m:
  STATE['mtime']=m;STATE['overrides']={};STATE['payload']=serial(*recalc())
 return STATE['payload']
def export_bytes():
 p=payload();df=pd.DataFrame([{'CLAVE':x['key'],'FECHA_PROCESO':x['process_date'][:10],'CLIENTE':x['client'],'FINCA':x['farm'],'DESTINO':x['destination'],'TARIMAS':x['pallets'],'PREENFRIO_ASIGNADO':x['assigned_origin'],'ORDEN':x['order'],'FECHA_CITA':x['appointment'][:10],'FECHA_INGRESO':x['start'],'FECHA_DESPACHO':x['end'],'ESTADO':x['status']} for x in p['data']]);b=io.BytesIO();df.to_excel(b,sheet_name='PLAN_AJUSTADO',index=False,engine='openpyxl');return b.getvalue()
def import_overrides(raw):
 df=pd.read_excel(io.BytesIO(raw),sheet_name='PLAN_AJUSTADO',engine='openpyxl');c={n:col(df,n) for n in ['CLAVE','PREENFRIO_ASIGNADO','ORDEN']};ov={}
 for _,r in df.iterrows():ov[str(r[c['CLAVE']]).strip()]={'origin':norm(r[c['PREENFRIO_ASIGNADO']]),'order':int(float(r[c['ORDEN']]))}
 STATE['overrides']=ov;STATE['payload']=serial(*recalc(ov));return STATE['payload']
class H(BaseHTTPRequestHandler):
 def sendx(self,b,t='text/plain; charset=utf-8',code=200,extra=None):
  self.send_response(code);self.send_header('Content-Type',t);self.send_header('Cache-Control','no-store');self.send_header('Content-Length',str(len(b)));[(self.send_header(k,v)) for k,v in (extra or {}).items()];self.end_headers();self.wfile.write(b)
 def jsonbody(self):return json.loads(self.rfile.read(int(self.headers.get('Content-Length','0'))) or b'{}')
 def do_GET(self):
  p=urlparse(self.path).path
  if p=='/':return self.sendx(HTML.read_bytes(),'text/html; charset=utf-8')
  if p=='/api/datos':return self.sendx(json.dumps(payload(),ensure_ascii=False).encode(),'application/json; charset=utf-8')
  if p=='/api/exportar':return self.sendx(export_bytes(),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',extra={'Content-Disposition':'attachment; filename="PLAN_AJUSTADO_V2.xlsx"'})
  return self.sendx(b'No encontrado',code=404)
 def do_POST(self):
  p=urlparse(self.path).path
  try:
   if p=='/api/replan':
    body=self.jsonbody();STATE['overrides']=body.get('overrides',{});STATE['payload']=serial(*recalc(STATE['overrides']));return self.sendx(json.dumps(STATE['payload'],ensure_ascii=False).encode(),'application/json; charset=utf-8')
   if p=='/api/importar':
    raw=self.rfile.read(int(self.headers.get('Content-Length','0')));result=import_overrides(raw);return self.sendx(json.dumps(result,ensure_ascii=False).encode(),'application/json; charset=utf-8')
  except Exception as e:return self.sendx(json.dumps({'ok':False,'error':str(e)},ensure_ascii=False).encode(),'application/json; charset=utf-8',500)
  self.sendx(b'No encontrado',code=404)
 def log_message(self,*a):pass
if __name__=='__main__':
 print(f'Planeador V2.1: http://{HOST}:{PORT}');s=ThreadingHTTPServer((HOST,PORT),H);threading.Timer(.6,lambda:webbrowser.open(f'http://{HOST}:{PORT}')).start()
 try:s.serve_forever()
 except KeyboardInterrupt:pass
