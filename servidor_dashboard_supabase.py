#!/usr/bin/env python3
"""Planeador Supabase con programación FIFO diaria y preenfrío parcial inmediato."""
from __future__ import annotations
import io,json,os,threading,time as clock
from datetime import datetime,timedelta
from http.server import ThreadingHTTPServer
from urllib.parse import urlparse
import pandas as pd
from supabase import create_client
import servidor_dashboard_v2 as core

HOST="0.0.0.0" if os.environ.get("PORT") else "127.0.0.1"
PORT=int(os.environ.get("PORT","8010")); CACHE_SECONDS=int(os.environ.get("CACHE_SECONDS","45"))
URL=os.environ.get("SUPABASE_URL"); KEY=os.environ.get("SUPABASE_SERVICE_ROLE_KEY") or os.environ.get("SUPABASE_SECRET_KEY")
if not URL or not KEY: raise RuntimeError("Faltan SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY")
DB=create_client(URL,KEY); LOCK=threading.RLock(); CACHE={"payload":None,"time":0.0}
CATALOGUE=core.read_catalogue(); LOCATIONS,PRECOOLERS=core.read_locations()

def parse_dt(v):
    d=pd.to_datetime(v,errors="coerce"); return None if pd.isna(d) else d.to_pydatetime()

def fetch_all(table,select="*",chunk=1000):
    rows=[]; start=0
    while True:
        data=DB.table(table).select(select).range(start,start+chunk-1).execute().data; rows.extend(data)
        if len(data)<chunk: return rows
        start+=chunk

def load_jobs_and_plan():
    source=fetch_all("procesos_fuente","*,procesos!inner(id,activo_fuente)"); plan_rows=fetch_all("procesos_planificacion","*")
    plans={r["proceso_id"]:r for r in plan_rows}; jobs=[]
    for index,original in enumerate(source,1):
        row=dict(original); meta=row.pop("procesos")
        if not meta.get("activo_fuente",True): continue
        destination=str(row.get("destino") or "").strip(); info=CATALOGUE.get(core.norm(destination))
        process_date,appointment=parse_dt(row.get("fecha_proceso")),parse_dt(row.get("fecha_cita"))
        try: pallets=float(row.get("pallets") or 0)
        except (TypeError,ValueError): pallets=0
        if not info or not process_date or not appointment or pallets<=0: continue
        farm=str(row.get("finca") or "").strip(); coord=LOCATIONS.get(core.norm(farm))
        distances={o:core.haversine_km(coord,p) for o,p in PRECOOLERS.items()}
        jobs.append({"key":meta["id"],"source":index,"farm":farm,"destination":destination,
                     "client":core.norm(row.get("cliente") or info["client"]),
                     "process_date":process_date.replace(hour=0,minute=0,second=0,microsecond=0),"appointment":appointment,
                     "transit_h":float(row.get("horas_transito") if row.get("horas_transito") is not None else info["transit_h"]),
                     "critical":int(row.get("nivel_critico") if row.get("nivel_critico") is not None else info["critical"]),
                     "pallets":pallets,"week":str(row.get("semana") or process_date.isocalendar().week),
                     "cycle_options":info["cycle"],"distance_options_km":distances,"location_found":coord is not None,
                     "ov":row.get("ov"),"sku":row.get("sku"),"lote":row.get("lote")})
    return jobs,plans

def fast_automatic(jobs):
    queues={"NELLY":[],"FORTALEZA":[]}; assignments={}
    ordered=sorted(jobs,key=lambda j:(core.release_time(j),core.deadline(j),j.get("critical",99),j["source"]))
    for job in ordered:
        best=None
        for origin in core.allowed(job):
            rows,cycles=core.pack(origin,queues[origin]+[job]); candidate=next(r for r in rows if r["key"]==job["key"])
            protected_late=sum(r["status"]=="NO CUMPLE" and r.get("critical",99)<=4 for r in rows)
            protected_delay=sum(max(0,-r["slack_h"]) for r in rows if r.get("critical",99)<=4)
            score=(protected_late,round(protected_delay,6),0 if candidate["status"]=="CUMPLE" else 1,max(0,-candidate["slack_h"]),len(cycles))
            if best is None or score<best[0]: best=(score,origin)
        if best:
            origin=best[1]; queues[origin].append(job)
            same_day=sum(x["process_date"].date()==job["process_date"].date() for x in queues[origin])
            assignments[job["key"]]={"origin":origin,"order":same_day}
    return assignments

def calculate():
    jobs,plans=load_jobs_and_plan(); automatic=fast_automatic(jobs); queues={"NELLY":[],"FORTALEZA":[]}; direct=[]
    for job in jobs:
        if job["key"] not in automatic: continue
        plan=plans.get(job["key"],{}); suggested=automatic[job["key"]]; permitted=core.allowed(job)
        assignment_type=plan.get("tipo_asignacion") or "AUTOMATICA"
        if assignment_type=="DIRECTO_SIN_PREENFRIO": origin="DIRECTO SIN PREENFRIO"
        else:
            if not permitted: continue
            origin=plan.get("preenfrio_manual") or suggested["origin"]
            if origin not in permitted: origin=suggested["origin"] if suggested["origin"] in permitted else permitted[0]
        try: order=int(plan.get("orden_manual") or suggested["order"])
        except (TypeError,ValueError): order=int(suggested["order"])
        item={**job,"manual_order":order,"suggested_origin":suggested["origin"],"suggested_order":suggested["order"]}
        if origin=="DIRECTO SIN PREENFRIO": direct.append(item)
        else: queues[origin].append(item)
    rows=[]; cycles=[]
    for origin in ("NELLY","FORTALEZA"):
        r,c=core.pack(origin,queues[origin]); rows.extend(r); cycles.extend(c)
    for order,job in enumerate(sorted(direct,key=lambda j:(core.release_time(j),core.deadline(j))),1):
        dispatch=core.release_time(job); arrival=dispatch+timedelta(hours=job["transit_h"]); slack=(job["appointment"]-arrival).total_seconds()/3600
        rows.append({**job,"assigned_origin":"DIRECTO SIN PREENFRIO","assigned_distance_km":None,"order":order,
                     "start":dispatch,"end":dispatch,"pre_cooling_end":dispatch,"dispatch":dispatch,"arrival":arrival,
                     "conservation_wait_h":0.0,"portions":[],"slack_h":slack,"status":"CUMPLE" if slack>=0 else "NO CUMPLE"})
    return rows,cycles,plans

def serialise(rows,cycles):
    def iso(v): return v.isoformat(timespec="minutes") if isinstance(v,datetime) else v
    rf=["key","farm","destination","client","process_date","appointment","transit_h","critical","pallets","week","assigned_origin","assigned_distance_km","distance_options_km","location_found","order","start","end","pre_cooling_end","dispatch","arrival","conservation_wait_h","portions","slack_h","status"]
    cf=["cycle_no","origin","process_date","start","end","duration_h","capacity","used_pallets","free_pallets","loads"]
    return {"ok":True,"source":"SUPABASE","version":"2.5","data":[{f:iso(r[f]) for f in rf} for r in rows],"cycles":[{f:iso(c[f]) for f in cf} for c in cycles]}

def payload(force=False):
    with LOCK:
        now=clock.monotonic()
        if not force and CACHE["payload"] is not None and now-CACHE["time"]<CACHE_SECONDS: return CACHE["payload"]
        rows,cycles,_=calculate(); CACHE.update(payload=serialise(rows,cycles),time=now); return CACHE["payload"]

def save_overrides(overrides):
    now=datetime.now().astimezone().isoformat()
    for process_id,change in overrides.items():
        origin=change.get("origin"); values={"orden_manual":int(change.get("order")) if change.get("order") else None,"modificado_por":"DASHBOARD","modificado_en":now}
        if origin=="DIRECTO SIN PREENFRIO": values.update(preenfrio_manual=None,tipo_asignacion="DIRECTO_SIN_PREENFRIO")
        elif origin in ("NELLY","FORTALEZA"): values.update(preenfrio_manual=origin,tipo_asignacion="MANUAL")
        DB.table("procesos_planificacion").update(values).eq("proceso_id",process_id).execute()

def export_bytes():
    p=payload(); df=pd.DataFrame([{"CLAVE":x["key"],"FECHA_PROCESO":x["process_date"][:10],"CLIENTE":x["client"],"FINCA":x["farm"],"DESTINO":x["destination"],"TARIMAS":x["pallets"],"PREENFRIO_ASIGNADO":x["assigned_origin"],"ORDEN":x["order"],"FECHA_CITA":x["appointment"],"INGRESO_PREENFRIO":x["start"],"FIN_PREENFRIO":x["pre_cooling_end"],"FECHA_DESPACHO":x["dispatch"],"ESTADO":x["status"]} for x in p["data"]])
    b=io.BytesIO(); df.to_excel(b,sheet_name="PLAN_AJUSTADO",index=False,engine="openpyxl"); return b.getvalue()

class Handler(core.H):
    def do_GET(self):
        path=urlparse(self.path).path
        try:
            if path=="/": return self.sendx(core.HTML.read_bytes(),"text/html; charset=utf-8")
            if path=="/api/datos": return self.sendx(json.dumps(payload(),ensure_ascii=False).encode(),"application/json; charset=utf-8")
            if path=="/api/exportar": return self.sendx(export_bytes(),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",extra={"Content-Disposition":'attachment; filename="PLAN_AJUSTADO.xlsx"'})
        except Exception as e: return self.sendx(json.dumps({"ok":False,"error":str(e)},ensure_ascii=False).encode(),"application/json; charset=utf-8",500)
        return self.sendx(b"No encontrado",code=404)
    def do_POST(self):
        try:
            raw=self.rfile.read(int(self.headers.get("Content-Length","0"))); path=urlparse(self.path).path
            if path=="/api/replan":
                save_overrides(json.loads(raw or b"{}").get("overrides",{})); result=payload(force=True)
                return self.sendx(json.dumps(result,ensure_ascii=False).encode(),"application/json; charset=utf-8")
            if path=="/api/importar": return self.sendx(json.dumps({"ok":False,"error":"La importación base se realiza con el importador incremental"},ensure_ascii=False).encode(),"application/json; charset=utf-8",409)
        except Exception as e: return self.sendx(json.dumps({"ok":False,"error":str(e)},ensure_ascii=False).encode(),"application/json; charset=utf-8",500)
        return self.sendx(b"No encontrado",code=404)

if __name__=="__main__":
    print(f"Planeador Supabase V2.5: http://{HOST}:{PORT}"); server=ThreadingHTTPServer((HOST,PORT),Handler)
    try: server.serve_forever()
    except KeyboardInterrupt: pass
    finally: server.server_close()
