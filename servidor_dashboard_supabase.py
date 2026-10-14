#!/usr/bin/env python3
"""Planeador Supabase con FIFO diario y preenfrío parcial inmediato."""
from __future__ import annotations

import io
import json
import os
import threading
import time as clock
from datetime import datetime, timedelta
from http.server import ThreadingHTTPServer
from urllib.parse import urlparse

import pandas as pd
from supabase import create_client

import servidor_dashboard_v2 as core
import cierres_runtime

HOST = "0.0.0.0" if os.environ.get("PORT") else "127.0.0.1"
PORT = int(os.environ.get("PORT", "8010"))
CACHE_SECONDS = int(os.environ.get("CACHE_SECONDS", "45"))
URL = os.environ.get("SUPABASE_URL")
KEY = os.environ.get("SUPABASE_SERVICE_ROLE_KEY") or os.environ.get("SUPABASE_SECRET_KEY")
if not URL or not KEY:
    raise RuntimeError("Faltan SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY")

DB = create_client(URL, KEY)
LOCK = threading.RLock()
CACHE = {"payload": None, "time": 0.0}
CATALOGUE = core.read_catalogue()
LOCATIONS, PRECOOLERS = core.read_locations()


def parse_dt(value):
    parsed = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(parsed) else parsed.to_pydatetime()


def fetch_all(table, select="*", chunk=1000):
    rows, start = [], 0
    while True:
        data = DB.table(table).select(select).range(start, start + chunk - 1).execute().data
        rows.extend(data)
        if len(data) < chunk:
            return rows
        start += chunk


def load_jobs_and_plan():
    source = fetch_all("procesos_fuente", "*,procesos!inner(id,activo_fuente)")
    plan_rows = fetch_all("procesos_planificacion", "*")
    plans = {row["proceso_id"]: row for row in plan_rows}
    jobs = []

    for index, original in enumerate(source, 1):
        row = dict(original)
        process_metadata = row.pop("procesos")
        if not process_metadata.get("activo_fuente", True):
            continue

        destination = str(row.get("destino") or "").strip()
        info = CATALOGUE.get(core.norm(destination))
        process_date = parse_dt(row.get("fecha_proceso"))
        appointment = parse_dt(row.get("fecha_cita"))
        try:
            pallets = float(row.get("pallets") or 0)
        except (TypeError, ValueError):
            pallets = 0
        if not info or not process_date or not appointment or pallets <= 0:
            continue

        farm = str(row.get("finca") or "").strip()
        coordinate = LOCATIONS.get(core.norm(farm))
        distances = {
            origin: core.haversine_km(coordinate, precooler)
            for origin, precooler in PRECOOLERS.items()
        }
        jobs.append(
            {
                "key": process_metadata["id"],
                "source": index,
                "farm": farm,
                "destination": destination,
                "client": core.norm(row.get("cliente") or info["client"]),
                "process_date": process_date.replace(hour=0, minute=0, second=0, microsecond=0),
                "appointment": appointment,
                "transit_h": float(
                    row.get("horas_transito")
                    if row.get("horas_transito") is not None
                    else info["transit_h"]
                ),
                "critical": int(
                    row.get("nivel_critico")
                    if row.get("nivel_critico") is not None
                    else info["critical"]
                ),
                "pallets": pallets,
                "week": str(row.get("semana") or process_date.isocalendar().week),
                "cycle_options": info["cycle"],
                "distance_options_km": distances,
                "location_found": coordinate is not None,
                "ov": row.get("ov"),
                "sku": row.get("sku"),
                "lote": row.get("lote"),
            }
        )
    return jobs, plans


def fast_automatic(jobs, initial_clocks=None):
    initial_clocks = initial_clocks or {"NELLY": None, "FORTALEZA": None}
    queues = {"NELLY": [], "FORTALEZA": []}
    assignments = {}
    ordered = sorted(
        jobs,
        key=lambda job: (
            core.release_time(job),
            job.get("critical", 99),
            core.deadline(job),
            job["source"],
        ),
    )

    for job in ordered:
        best = None
        for origin in core.allowed(job):
            rows, cycles = core.pack(origin, queues[origin] + [job], initial_clocks.get(origin))
            candidate = next(row for row in rows if row["key"] == job["key"])
            protected_late = sum(
                row["status"] == "NO CUMPLE" and row.get("critical", 99) <= 4
                for row in rows
            )
            protected_delay = sum(
                max(0, -row["slack_h"])
                for row in rows
                if row.get("critical", 99) <= 4
            )
            score = (
                protected_late,
                round(protected_delay, 6),
                0 if candidate["status"] == "CUMPLE" else 1,
                max(0, -candidate["slack_h"]),
                len(cycles),
            )
            if best is None or score < best[0]:
                best = (score, origin)

        if best:
            origin = best[1]
            queues[origin].append(job)
            same_day_count = sum(
                item["process_date"].date() == job["process_date"].date()
                for item in queues[origin]
            )
            assignments[job["key"]] = {"origin": origin, "order": same_day_count}
    return assignments


def historical_rows(closures):
    if not closures:
        return []
    closure_by_id = {row["id"]: row for row in closures}
    details = fetch_all("cierres_semanales_detalle", "*")
    result = []
    for position, detail in enumerate(details, 1):
        if detail.get("cierre_id") not in closure_by_id:
            continue
        planned_origin = detail.get("preenfrio_planeado") or "SIN REGISTRO"
        real_origin = detail.get("preenfrio_real")
        planned_start = parse_dt(detail.get("ingreso_planeado")); planned_end = parse_dt(detail.get("fin_planeado"))
        real_start = parse_dt(detail.get("ingreso_real")); real_end = parse_dt(detail.get("fin_real"))
        planned_dispatch = parse_dt(detail.get("despacho_planeado")); real_dispatch = parse_dt(detail.get("despacho_real"))
        result.append({
            "key": detail.get("proceso_id") or detail.get("snapshot_key"), "source": position,
            "farm": detail.get("finca") or "", "destination": detail.get("destino") or "",
            "client": detail.get("cliente") or "", "sku": detail.get("sku"), "lote": detail.get("lote"),
            "process_date": parse_dt(detail.get("fecha_proceso")), "appointment": parse_dt(detail.get("fecha_cita")),
            "transit_h": 0.0, "critical": 99, "pallets": float(detail.get("pallets") or 0),
            "week": str(detail.get("semana")), "assigned_origin": real_origin or planned_origin,
            "assigned_distance_km": None, "distance_options_km": {}, "location_found": True,
            "order": int(detail.get("orden_planeado") or position), "start": real_start or planned_start,
            "end": real_end or planned_end, "pre_cooling_end": real_end or planned_end,
            "dispatch": real_dispatch or planned_dispatch or real_end or planned_end,
            "arrival": parse_dt(detail.get("llegada_planeada")), "conservation_wait_h": None,
            "portions": [], "slack_h": None,
            "status": detail.get("cumplimiento_real") or detail.get("cumplimiento_planeado") or detail.get("estado_cierre") or "CERRADA",
            "read_only": True, "closed": True, "closed_year": int(detail.get("anio")), "closed_week": int(detail.get("semana")),
            "planned_origin": planned_origin, "planned_start": planned_start, "planned_end": planned_end,
            "planned_dispatch": planned_dispatch, "planned_status": detail.get("cumplimiento_planeado"),
            "real_origin": real_origin, "real_start": real_start, "real_end": real_end,
            "real_dispatch": real_dispatch, "real_status": detail.get("cumplimiento_real"),
            "has_real": bool(detail.get("tiene_real")), "time_source": detail.get("fuente_tiempo") or "PLANEACION_SIN_REAL",
            "closure_status": detail.get("estado_cierre"), "observations": detail.get("observaciones"),
        })
    return result


def calculate(transient_overrides=None):
    all_jobs, plans = load_jobs_and_plan()
    closures = cierres_runtime.obtener_cierres(DB)
    latest = closures[-1] if closures else None
    closed_keys = cierres_runtime.claves_cerradas(closures)
    jobs = [j for j in all_jobs if cierres_runtime.clave_semana_job(j) not in closed_keys]
    clocks = cierres_runtime.relojes_iniciales(latest)
    transient_overrides = transient_overrides or {}
    automatic_assignments = fast_automatic(jobs, clocks)
    queues = {"NELLY": [], "FORTALEZA": []}; direct = []
    for job in jobs:
        if job["key"] not in automatic_assignments: continue
        plan = dict(plans.get(job["key"], {})); change = transient_overrides.get(job["key"])
        if change:
            if change.get("order") is not None: plan["orden_manual"] = int(change["order"])
            if change.get("origin") == "DIRECTO SIN PREENFRIO": plan.update(preenfrio_manual=None,tipo_asignacion="DIRECTO_SIN_PREENFRIO")
            elif change.get("origin") in ("NELLY","FORTALEZA"): plan.update(preenfrio_manual=change["origin"],tipo_asignacion="MANUAL")
        suggested=automatic_assignments[job["key"]]; typ=plan.get("tipo_asignacion") or "AUTOMATICA"
        if typ=="DIRECTO_SIN_PREENFRIO": origin="DIRECTO SIN PREENFRIO"
        else:
            manual=typ=="MANUAL" and bool(plan.get("preenfrio_manual")); permitted=core.manual_allowed(job) if manual else core.allowed(job)
            if not permitted: continue
            origin=plan.get("preenfrio_manual") if manual else suggested["origin"]
            if origin not in permitted: origin=suggested["origin"] if suggested["origin"] in permitted else permitted[0]
        try: order=int(plan.get("orden_manual") or suggested["order"])
        except: order=int(suggested["order"])
        item={**job,"manual_order":order,"suggested_origin":suggested["origin"],"suggested_order":suggested["order"],"read_only":False,"closed":False}
        (direct if origin=="DIRECTO SIN PREENFRIO" else queues[origin]).append(item)
    rows=[]; cycles=[]
    for origin in ("NELLY","FORTALEZA"):
        packed, cyc=core.pack(origin,queues[origin],clocks.get(origin)); rows.extend(packed); cycles.extend(cyc)
    for order,job in enumerate(sorted(direct,key=lambda x:(core.release_time(x),core.deadline(x))),1):
        dispatch=core.release_time(job); arrival=dispatch+timedelta(hours=job["transit_h"]); slack=(job["appointment"]-arrival).total_seconds()/3600
        rows.append({**job,"assigned_origin":"DIRECTO SIN PREENFRIO","assigned_distance_km":None,"order":order,"start":dispatch,"end":dispatch,"pre_cooling_end":dispatch,"dispatch":dispatch,"arrival":arrival,"conservation_wait_h":0.0,"portions":[],"slack_h":slack,"status":"CUMPLE" if slack>=0 else "NO CUMPLE"})
    return rows,cycles,plans,historical_rows(closures),closures

def json_value(value):
    """Conversión recursiva a tipos compatibles con JSON."""
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().isoformat(timespec="minutes")
    if isinstance(value, datetime):
        return value.isoformat(timespec="minutes")
    if isinstance(value, dict):
        return {key: json_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_value(item) for item in value]
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    return value


def serialise(rows, cycles, historical, closures):
    row_fields=["key","farm","destination","client","process_date","appointment","transit_h","critical","pallets","week","assigned_origin","assigned_distance_km","distance_options_km","location_found","order","start","end","pre_cooling_end","dispatch","arrival","conservation_wait_h","portions","slack_h","status","read_only","closed","closed_year","closed_week","planned_origin","planned_start","planned_end","planned_dispatch","planned_status","real_origin","real_start","real_end","real_dispatch","real_status","has_real","time_source","closure_status","observations"]
    cycle_fields=["cycle_no","origin","process_date","start","end","duration_h","capacity","used_pallets","free_pallets","loads"]
    enc=lambda src:[{f:json_value(r.get(f)) for f in row_fields} for r in src]
    return {"ok":True,"source":"SUPABASE","version":"2.9.0","data":enc(rows),"historical_data":enc(historical),"closures":[json_value(cierres_runtime.cierre_publico(c)) for c in closures],"cycles":[{f:json_value(c.get(f)) for f in cycle_fields} for c in cycles]}

def payload(force=False, transient_overrides=None):
    with LOCK:
        now = clock.monotonic()
        if (
            not force
            and CACHE["payload"] is not None
            and now - CACHE["time"] < CACHE_SECONDS
        ):
            return CACHE["payload"]
        rows, cycles, _, historical, closures = calculate(transient_overrides)
        CACHE.update(payload=serialise(rows, cycles, historical, closures), time=now)
        return CACHE["payload"]


def save_overrides(overrides):
    if not overrides: return
    closed= cierres_runtime.claves_cerradas(cierres_runtime.obtener_cierres(DB)); ids=list(overrides); rows=[]
    for i in range(0,len(ids),100): rows.extend(DB.table("procesos_fuente").select("proceso_id,fecha_proceso").in_("proceso_id",ids[i:i+100]).execute().data or [])
    by_id={r["proceso_id"]:r for r in rows}
    for pid in ids:
        dt=parse_dt((by_id.get(pid) or {}).get("fecha_proceso"))
        if dt and (dt.isocalendar().year,dt.isocalendar().week) in closed: raise PermissionError("Semana cerrada en modo solo consulta")
    modified_at=datetime.now().astimezone().isoformat()
    for pid,change in overrides.items():
        origin=change.get("origin"); values={"orden_manual":int(change.get("order")) if change.get("order") else None,"modificado_por":"DASHBOARD","modificado_en":modified_at}
        if origin=="DIRECTO SIN PREENFRIO": values.update(preenfrio_manual=None,tipo_asignacion="DIRECTO_SIN_PREENFRIO")
        elif origin in ("NELLY","FORTALEZA"): values.update(preenfrio_manual=origin,tipo_asignacion="MANUAL")
        DB.table("procesos_planificacion").update(values).eq("proceso_id",pid).execute()

def export_bytes():
    current = payload()
    frame = pd.DataFrame(
        [
            {
                "CLAVE": item["key"],
                "FECHA_PROCESO": item["process_date"][:10],
                "CLIENTE": item["client"],
                "FINCA": item["farm"],
                "DESTINO": item["destination"],
                "TARIMAS": item["pallets"],
                "PREENFRIO_ASIGNADO": item["assigned_origin"],
                "ORDEN": item["order"],
                "FECHA_CITA": item["appointment"],
                "INGRESO_PREENFRIO": item["start"],
                "FIN_PREENFRIO": item["pre_cooling_end"],
                "FECHA_DESPACHO": item["dispatch"],
                "ESTADO": item["status"],
            }
            for item in current["data"]
        ]
    )
    buffer = io.BytesIO()
    frame.to_excel(buffer, sheet_name="PLAN_AJUSTADO", index=False, engine="openpyxl")
    return buffer.getvalue()


class Handler(core.H):
    def do_GET(self):
        path = urlparse(self.path).path
        try:
            if path == "/":
                return self.sendx(core.HTML.read_bytes(), "text/html; charset=utf-8")
            if path == "/imagen_portada.jpg":
                image_path = core.BASE / "imagen_portada.jpg"
                if not image_path.exists():
                    return self.sendx(b"Imagen no encontrada", code=404)
                return self.sendx(image_path.read_bytes(), "image/jpeg")
            if path == "/fondo_body.jpg":
                image_path = core.BASE / "fondo_body.jpg"
                if not image_path.exists():
                    return self.sendx(b"Imagen no encontrada", code=404)
                return self.sendx(image_path.read_bytes(), "image/jpeg")
            if path == "/api/datos":
                body = json.dumps(payload(), ensure_ascii=False, default=str).encode()
                return self.sendx(body, "application/json; charset=utf-8")
            if path == "/api/exportar":
                return self.sendx(
                    export_bytes(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    extra={"Content-Disposition": 'attachment; filename="PLAN_AJUSTADO.xlsx"'},
                )
        except Exception as exc:
            body = json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False).encode()
            return self.sendx(body, "application/json; charset=utf-8", 500)
        return self.sendx(b"No encontrado", code=404)

    def do_POST(self):
        try:
            raw = self.rfile.read(int(self.headers.get("Content-Length", "0")))
            path = urlparse(self.path).path
            if path == "/api/replan":
                overrides = json.loads(raw or b"{}").get("overrides", {})
                save_overrides(overrides)
                # Reflejo inmediato del nuevo orden, sin depender de una segunda
                # lectura a Supabase dentro de la misma solicitud.
                result = payload(force=True, transient_overrides=overrides)
                body = json.dumps(result, ensure_ascii=False, default=str).encode()
                return self.sendx(body, "application/json; charset=utf-8")
            if path == "/api/importar":
                body = json.dumps(
                    {
                        "ok": False,
                        "error": "La importación base se realiza con el importador incremental",
                    },
                    ensure_ascii=False,
                ).encode()
                return self.sendx(body, "application/json; charset=utf-8", 409)
        except Exception as exc:
            body = json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False).encode()
            return self.sendx(body, "application/json; charset=utf-8", 500)
        return self.sendx(b"No encontrado", code=404)


if __name__ == "__main__":
    print(f"Planeador Supabase V2.9.0: http://{HOST}:{PORT}")
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
