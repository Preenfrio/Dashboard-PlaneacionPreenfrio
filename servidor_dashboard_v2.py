#!/usr/bin/env python3
from __future__ import annotations

import io, json, math, os, re, threading, webbrowser
from datetime import datetime, timedelta, time
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse
import pandas as pd

BASE = Path(__file__).resolve().parent
EXCEL = BASE / "PLANEACION_SEMANAL 1.xlsx"
HTML = BASE / "dashboard_preenfrio_v2.html"
HOST = "0.0.0.0" if os.environ.get("PORT") else "127.0.0.1"
PORT = int(os.environ.get("PORT", "8010"))
CAP = {"NELLY": 12.0, "FORTALEZA": 10.0}
START_HOUR = 19
GAP_MIN = 30
STATE = {"overrides": {}, "payload": None, "mtime": None}


def norm(v): return " ".join(str(v if v is not None else "").strip().upper().split())

def col(df, name, required=True):
    wanted = norm(name)
    for c in df.columns:
        if norm(c) == wanted: return c
    if required: raise ValueError(f"Falta columna: {name}")
    return None

def as_dt(v):
    d = pd.to_datetime(v, dayfirst=True, errors="coerce")
    return None if pd.isna(d) else d.to_pydatetime()

def hours(v):
    if v is None or (isinstance(v, float) and pd.isna(v)) or v == "": return 0.0
    if isinstance(v, (datetime, time)): return v.hour + v.minute/60 + v.second/3600
    if isinstance(v, timedelta): return v.total_seconds()/3600
    if isinstance(v, (int, float)):
        x = float(v); return x*24 if 0 <= x < 1.5 else x
    p = str(v).strip().split(":")
    try: return float(p[0]) + (float(p[1])/60 if len(p)>1 else 0) + (float(p[2])/3600 if len(p)>2 else 0)
    except (TypeError, ValueError): return 0.0

def parse_coordinate(v):
    if v is None or (isinstance(v,float) and math.isnan(v)): return None
    text = str(v).strip().upper().replace("–","-").replace("—","-")
    dms = re.findall(r'(\d{1,3})\s*°\s*(\d{1,2})\s*[\'’]\s*(\d+(?:\.\d+)?)\s*[\"”]?\s*([NSEW])', text)
    if len(dms) >= 2:
        out=[]
        for dg,mn,sc,h in dms[:2]:
            n=float(dg)+float(mn)/60+float(sc)/3600
            out.append(-n if h in ("S","W") else n)
        return tuple(out)
    nums=re.findall(r"[-+]?\d{1,3}(?:\.\d+)",text)
    if len(nums)>=2:
        lat,lon=float(nums[0]),float(nums[1])
        if "W" in text and lon>0: lon=-lon
        if "S" in text and lat>0: lat=-lat
        if -90<=lat<=90 and -180<=lon<=180: return lat,lon
    return None

def haversine_km(a,b):
    if not a or not b: return None
    lat1,lon1=map(math.radians,a); lat2,lon2=map(math.radians,b)
    dlat,dlon=lat2-lat1,lon2-lon1
    h=math.sin(dlat/2)**2+math.cos(lat1)*math.cos(lat2)*math.sin(dlon/2)**2
    return 6371.0088*2*math.asin(math.sqrt(h))

def read_locations():
    df=pd.read_excel(EXCEL,sheet_name="UBICACIONES",engine="openpyxl")
    nc,cc=col(df,"Trazabilidad de finca"),col(df,"Coordenadas")
    loc={}
    for _,r in df.iterrows():
        name,coord=norm(r[nc]),parse_coordinate(r[cc])
        if name and coord: loc[name]=coord
    return loc,{"NELLY":loc.get("PREENFRIO NELLY"),"FORTALEZA":loc.get("PREENFRIO FORTALEZA")}

def read_catalogue():
    df=pd.read_excel(EXCEL,sheet_name="CATALOGO",engine="openpyxl")
    names=["CLIENTE","DESTINO","TRÁNSITO (HRS)","CICLO HORAS FORTALEZA","CICLO HORAS NELLY","CRITICIDAD"]
    c={n:col(df,n) for n in names}; out={}
    for _,r in df.iterrows():
        d=norm(r[c["DESTINO"]])
        if not d: continue
        try: critical=int(float(r[c["CRITICIDAD"]]))
        except (TypeError,ValueError): critical=99
        out[d]={"client":norm(r[c["CLIENTE"]]),"transit_h":hours(r[c["TRÁNSITO (HRS)"]]),"critical":critical,
                "cycle":{"NELLY":hours(r[c["CICLO HORAS NELLY"]]),"FORTALEZA":hours(r[c["CICLO HORAS FORTALEZA"]])}}
    return out

def read_jobs():
    catalogue=read_catalogue(); locations,precoolers=read_locations()
    df=pd.read_excel(EXCEL,sheet_name="PLANEACION",engine="openpyxl")
    c={n:col(df,n) for n in ["Finca","Destino","Pallets","Fecha de proceso","FECHA CITA"]}
    for n in ["SEMANA","HORAS TRANSITO","NIVEL CRITICO"]: c[n]=col(df,n,False)
    jobs=[]
    for i,r in df.iterrows():
        farm,destination=str(r[c["Finca"]]).strip(),str(r[c["Destino"]]).strip(); info=catalogue.get(norm(destination))
        process_date,appointment=as_dt(r[c["Fecha de proceso"]]),as_dt(r[c["FECHA CITA"]])
        try: pallets=float(r[c["Pallets"]])
        except (TypeError,ValueError): pallets=0
        if not info or not process_date or not appointment or pallets<=0: continue
        transit=hours(r[c["HORAS TRANSITO"]]) if c["HORAS TRANSITO"] else info["transit_h"]
        try: critical=int(float(r[c["NIVEL CRITICO"]])) if c["NIVEL CRITICO"] else info["critical"]
        except (TypeError,ValueError): critical=info["critical"]
        try: week=str(int(float(r[c["SEMANA"]]))) if c["SEMANA"] else str(process_date.isocalendar().week)
        except (TypeError,ValueError): week=str(process_date.isocalendar().week)
        coord=locations.get(norm(farm)); distances={o:haversine_km(coord,p) for o,p in precoolers.items()}
        key=f"{norm(farm)}|{norm(destination)}|{process_date.date()}|{appointment.date()}|{info['client']}|{i+2}"
        jobs.append({"key":key,"source":i+2,"farm":farm,"destination":destination,"client":info["client"],
                     "process_date":process_date.replace(hour=0,minute=0,second=0,microsecond=0),"appointment":appointment,
                     "transit_h":transit,"critical":critical,"pallets":pallets,"week":week,"cycle_options":info["cycle"],
                     "distance_options_km":distances,"location_found":coord is not None})
    return jobs

def allowed(job):
    farm,client=norm(job.get("farm")),norm(job.get("client")); options=job.get("cycle_options") or {}
    enabled={o:hours(options.get(o,0))>0 for o in CAP}
    if farm.startswith("A01 003") or "DOÑA NELLY" in farm or "DONA NELLY" in farm: return ["NELLY"] if enabled["NELLY"] else []
    if farm.startswith("A01 002") or farm=="FORTALEZA" or "A01 002 FORTALEZA" in farm: return ["FORTALEZA"] if enabled["FORTALEZA"] else []
    if client=="SMART": return ["NELLY"] if enabled["NELLY"] else []
    return [o for o in ("NELLY","FORTALEZA") if enabled[o]]

def release_time(job): return datetime.combine(job["process_date"].date(),time(START_HOUR))

def deadline(job): return job["appointment"]-timedelta(hours=job["transit_h"])

def daily_priority(job,origin):
    distance=job.get("distance_options_km",{}).get(origin)
    return (deadline(job),job.get("critical",99),distance if distance is not None else 1_000_000.0,job.get("manual_order",9999),job.get("source",999999))

def pack(origin,sequence):
    """FIFO estricto por día. Dentro del día aplica cita, criticidad, distancia y orden manual.
    Cada porción disponible entra cuanto antes; los remanentes continúan en ciclos posteriores.
    """
    if not sequence: return [],[]
    capacity=CAP[origin]; remaining={j["key"]:float(j["pallets"]) for j in sequence}
    first,last,portions,cycles={},{},{},[]; pending=list(sequence); current=min(release_time(j) for j in pending)
    while pending:
        available=[j for j in pending if release_time(j)<=current and remaining[j["key"]]>1e-9]
        if not available:
            current=min(release_time(j) for j in pending if remaining[j["key"]]>1e-9)
            available=[j for j in pending if release_time(j)<=current and remaining[j["key"]]>1e-9]
        oldest=min(j["process_date"].date() for j in available)
        day_jobs=[j for j in available if j["process_date"].date()==oldest]
        day_jobs.sort(key=lambda j:daily_priority(j,origin))
        space,loads,duration=capacity,[],0.0
        for job in day_jobs:
            if space<=1e-9: break
            qty=min(space,remaining[job["key"]])
            if qty<=1e-9: continue
            loads.append({"key":job["key"],"destination":job["destination"],"farm":job.get("farm",""),"pallets":qty,
                          "production_date":job["process_date"].date().isoformat()})
            first.setdefault(job["key"],current); portions.setdefault(job["key"],[]).append({"start":current,"pallets":qty})
            remaining[job["key"]]-=qty; space-=qty; duration=max(duration,hours(job["cycle_options"].get(origin,0)))
        if not loads: raise ValueError(f"No se pudo formar ciclo para {origin} el {oldest}")
        end=current+timedelta(hours=duration)
        for load in loads: portions[load["key"]][-1]["end"]=end
        cycles.append({"cycle_no":len(cycles)+1,"origin":origin,"process_date":datetime.combine(oldest,time()),
                       "start":current,"end":end,"duration_h":duration,"capacity":capacity,
                       "used_pallets":capacity-space,"free_pallets":space,"loads":loads})
        done=[]
        for load in loads:
            if remaining[load["key"]]<=1e-9: last[load["key"]]=end; done.append(load["key"])
        pending=[j for j in pending if j["key"] not in done]
        current=end+timedelta(minutes=GAP_MIN)
    rows=[]; counters={}
    for job in sequence:
        finish=last[job["key"]]; dispatch=finish; arrival=dispatch+timedelta(hours=job["transit_h"])
        slack=(job["appointment"]-arrival).total_seconds()/3600; day=(job["process_date"].date(),origin)
        counters[day]=counters.get(day,0)+1
        rows.append({**job,"assigned_origin":origin,"assigned_distance_km":job["distance_options_km"].get(origin),
                     "order":counters[day],"start":first[job["key"]],"end":finish,"pre_cooling_end":finish,
                     "dispatch":dispatch,"arrival":arrival,"conservation_wait_h":0.0,"portions":portions[job["key"]],
                     "slack_h":slack,"status":"CUMPLE" if slack>=0 else "NO CUMPLE"})
    return rows,cycles

def automatic(jobs):
    queues={"NELLY":[],"FORTALEZA":[]}; assignments={}
    for job in sorted(jobs,key=lambda j:(release_time(j),deadline(j),j.get("critical",99),j["source"])):
        best=None
        for origin in allowed(job):
            rows,cycles=pack(origin,queues[origin]+[job]); candidate=next(r for r in rows if r["key"]==job["key"])
            score=(0 if candidate["status"]=="CUMPLE" else 1,max(0,-candidate["slack_h"]),len(cycles))
            if best is None or score<best[0]: best=(score,origin)
        if best:
            origin=best[1]; queues[origin].append(job)
            assignments[job["key"]]={"origin":origin,"order":1+sum(x["process_date"].date()==job["process_date"].date() for x in queues[origin][:-1])}
    return assignments

def recalc(overrides=None):
    jobs=read_jobs(); auto=automatic(jobs); ov=overrides if overrides is not None else STATE["overrides"]
    queues={"NELLY":[],"FORTALEZA":[]}; direct=[]
    for job in jobs:
        if job["key"] not in auto: continue
        chosen=ov.get(job["key"],auto[job["key"]]); origin=chosen.get("origin") or auto[job["key"]]["origin"]
        copy={**job,"manual_order":int(chosen.get("order") or auto[job["key"]]["order"])}
        if origin=="DIRECTO SIN PREENFRIO": direct.append(copy); continue
        permitted=allowed(job)
        if origin not in permitted: origin=auto[job["key"]]["origin"] if auto[job["key"]]["origin"] in permitted else permitted[0]
        queues[origin].append(copy)
    rows,cycles=[],[]
    for origin in ("NELLY","FORTALEZA"):
        r,c=pack(origin,queues[origin]); rows.extend(r); cycles.extend(c)
    for order,job in enumerate(sorted(direct,key=lambda j:(release_time(j),deadline(j))),1):
        dispatch=release_time(job); arrival=dispatch+timedelta(hours=job["transit_h"]); slack=(job["appointment"]-arrival).total_seconds()/3600
        rows.append({**job,"assigned_origin":"DIRECTO SIN PREENFRIO","assigned_distance_km":None,"order":order,
                     "start":dispatch,"end":dispatch,"pre_cooling_end":dispatch,"dispatch":dispatch,"arrival":arrival,
                     "conservation_wait_h":0.0,"portions":[],"slack_h":slack,"status":"CUMPLE" if slack>=0 else "NO CUMPLE"})
    return rows,cycles

def serial(rows,cycles):
    def iso(v): return v.isoformat(timespec="minutes") if isinstance(v,datetime) else v
    row_fields=["key","farm","destination","client","process_date","appointment","transit_h","critical","pallets","week","assigned_origin","assigned_distance_km","distance_options_km","location_found","order","start","end","pre_cooling_end","dispatch","arrival","conservation_wait_h","portions","slack_h","status"]
    cycle_fields=["cycle_no","origin","process_date","start","end","duration_h","capacity","used_pallets","free_pallets","loads"]
    return {"ok":True,"version":"2.5","excel":EXCEL.name,"modified":datetime.fromtimestamp(EXCEL.stat().st_mtime).isoformat(timespec="seconds"),
            "data":[{f:iso(r[f]) for f in row_fields} for r in rows],"cycles":[{f:iso(c[f]) for f in cycle_fields} for c in cycles]}

def payload():
    m=EXCEL.stat().st_mtime
    if STATE["payload"] is None or STATE["mtime"]!=m:
        STATE.update(mtime=m,overrides={}); STATE["payload"]=serial(*recalc())
    return STATE["payload"]

def export_bytes():
    p=payload(); df=pd.DataFrame([{"CLAVE":x["key"],"FECHA_PROCESO":x["process_date"][:10],"CLIENTE":x["client"],"FINCA":x["farm"],"DESTINO":x["destination"],"TARIMAS":x["pallets"],"PREENFRIO_ASIGNADO":x["assigned_origin"],"ORDEN":x["order"],"FECHA_CITA":x["appointment"],"INGRESO_PREENFRIO":x["start"],"FIN_PREENFRIO":x["pre_cooling_end"],"FECHA_DESPACHO":x["dispatch"],"ESTADO":x["status"]} for x in p["data"]])
    b=io.BytesIO(); df.to_excel(b,sheet_name="PLAN_AJUSTADO",index=False,engine="openpyxl"); return b.getvalue()

class H(BaseHTTPRequestHandler):
    def sendx(self,body,content_type="text/plain; charset=utf-8",code=200,extra=None):
        self.send_response(code); self.send_header("Content-Type",content_type); self.send_header("Cache-Control","no-store"); self.send_header("Content-Length",str(len(body)))
        for k,v in (extra or {}).items(): self.send_header(k,v)
        self.end_headers(); self.wfile.write(body)
    def do_GET(self):
        path=urlparse(self.path).path
        try:
            if path=="/": return self.sendx(HTML.read_bytes(),"text/html; charset=utf-8")
            if path=="/api/datos": return self.sendx(json.dumps(payload(),ensure_ascii=False).encode(),"application/json; charset=utf-8")
            if path=="/api/exportar": return self.sendx(export_bytes(),"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",extra={"Content-Disposition":'attachment; filename="PLAN_AJUSTADO_V2.xlsx"'})
        except Exception as e: return self.sendx(json.dumps({"ok":False,"error":str(e)},ensure_ascii=False).encode(),"application/json; charset=utf-8",500)
        return self.sendx(b"No encontrado",code=404)
    def do_POST(self):
        try:
            raw=self.rfile.read(int(self.headers.get("Content-Length","0")))
            if urlparse(self.path).path=="/api/replan":
                STATE["overrides"]=json.loads(raw or b"{}").get("overrides",{}); STATE["payload"]=serial(*recalc(STATE["overrides"]))
                return self.sendx(json.dumps(STATE["payload"],ensure_ascii=False).encode(),"application/json; charset=utf-8")
        except Exception as e: return self.sendx(json.dumps({"ok":False,"error":str(e)},ensure_ascii=False).encode(),"application/json; charset=utf-8",500)
        return self.sendx(b"No encontrado",code=404)
    def log_message(self,*_): pass

if __name__=="__main__":
    print(f"Planeador V2.5: http://{HOST}:{PORT}"); server=ThreadingHTTPServer((HOST,PORT),H)
    if HOST=="127.0.0.1": threading.Timer(.6,lambda:webbrowser.open(f"http://{HOST}:{PORT}")).start()
    try: server.serve_forever()
    except KeyboardInterrupt: pass
    finally: server.server_close()
