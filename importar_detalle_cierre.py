#!/usr/bin/env python3
"""Crea el snapshot histórico Planeación vs Real para una semana cerrada.

No recalcula semanas cerradas en el dashboard. Conserva la planeación como
base y superpone Movimientos cuando existe una coincidencia real.
"""
from __future__ import annotations

import argparse
import math
import os
import re
import unicodedata
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
from supabase import create_client

import servidor_dashboard_v2 as core


def norm(value):
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return " ".join(text.strip().upper().split())


def headers(frame):
    return {norm(col): col for col in frame.columns}


def pick(columns, *names, required=False):
    for name in names:
        if norm(name) in columns:
            return columns[norm(name)]
    if required:
        raise ValueError(f"Falta columna requerida: {names[0]}")
    return None


def as_datetime(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, (int, float)):
        return datetime(1899, 12, 30) + timedelta(days=float(value))
    parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
    return None if pd.isna(parsed) else parsed.to_pydatetime()


def combine_date_time(date_value, time_value):
    date = as_datetime(date_value)
    if not date:
        return None
    if time_value is None or (isinstance(time_value, float) and math.isnan(time_value)):
        return None
    if isinstance(time_value, (int, float)):
        seconds = round((float(time_value) % 1) * 86400)
        return datetime.combine(date.date(), datetime.min.time()) + timedelta(seconds=seconds)
    parsed = pd.to_datetime(str(time_value), errors="coerce")
    if pd.isna(parsed):
        return None
    return datetime.combine(date.date(), parsed.to_pydatetime().time())


def hours(value):
    return core.hours(value)


def iso(value):
    return value.isoformat(timespec="minutes") if value else None


def base_key(row):
    date = row["process_date"].strftime("%Y%m%d")
    return "|".join((norm(row["sku"]), norm(row["lote"]), norm(row["farm"]), norm(row["destination"]), date))


def normalize_origin(value, destination=""):
    text = norm(value)
    dest = norm(destination)
    if "DIRECTO" in text or dest in {"CEDA", "JAPON", "JAPÓN"}:
        return "DIRECTO SIN PREENFRIO"
    if "NELLY" in text:
        return "NELLY"
    if "FORTALEZA" in text:
        return "FORTALEZA"
    return None


def fetch_all(db, table, select="*", chunk=1000):
    rows, start = [], 0
    while True:
        batch = db.table(table).select(select).range(start, start + chunk - 1).execute().data or []
        rows.extend(batch)
        if len(batch) < chunk:
            return rows
        start += chunk


def build_jobs(db, year, week):
    source = fetch_all(db, "procesos_fuente", "*,procesos!inner(id,activo_fuente)")
    plans = {row["proceso_id"]: row for row in fetch_all(db, "procesos_planificacion")}
    catalogue = core.read_catalogue()
    locations, precoolers = core.read_locations()
    jobs = []
    for index, original in enumerate(source, 1):
        row = dict(original)
        process = row.pop("procesos")
        if not process.get("activo_fuente", True):
            continue
        process_date = as_datetime(row.get("fecha_proceso"))
        appointment = as_datetime(row.get("fecha_cita"))
        if not process_date or process_date.isocalendar()[:2] != (year, week):
            continue
        destination = str(row.get("destino") or "").strip()
        info = catalogue.get(core.norm(destination))
        if not info or not appointment:
            continue
        try:
            pallets = float(row.get("pallets") or 0)
        except (TypeError, ValueError):
            continue
        if pallets <= 0:
            continue
        farm = str(row.get("finca") or "").strip()
        coordinate = locations.get(core.norm(farm))
        distances = {origin: core.haversine_km(coordinate, precooler) for origin, precooler in precoolers.items()}
        jobs.append({
            "key": process["id"], "source": index, "farm": farm,
            "destination": destination, "client": core.norm(row.get("cliente") or info["client"]),
            "process_date": process_date.replace(hour=0, minute=0, second=0, microsecond=0),
            "appointment": appointment, "transit_h": float(row.get("horas_transito") or info["transit_h"]),
            "critical": int(row.get("nivel_critico") or info["critical"]), "pallets": pallets,
            "week": str(week), "cycle_options": info["cycle"], "distance_options_km": distances,
            "location_found": coordinate is not None, "sku": row.get("sku"), "lote": row.get("lote"),
            "cantidad": row.get("cantidad"), "plan": plans.get(process["id"], {}),
        })
    return jobs


def calculate_week_plan(jobs):
    queues = {"NELLY": [], "FORTALEZA": []}
    direct = []
    ordered = sorted(jobs, key=lambda j: (core.release_time(j), j.get("critical", 99), core.deadline(j), j["source"]))
    counters = defaultdict(int)
    for job in ordered:
        plan = job.get("plan", {})
        assignment_type = plan.get("tipo_asignacion") or "AUTOMATICA"
        if assignment_type == "DIRECTO_SIN_PREENFRIO":
            direct.append({**job, "manual_order": 0})
            continue
        manual = assignment_type == "MANUAL" and plan.get("preenfrio_manual")
        allowed = core.manual_allowed(job) if manual else core.allowed(job)
        if not allowed:
            continue
        if manual in allowed:
            origin = manual
        else:
            # Conserva reglas automáticas y elige la opción que cumple mejor.
            candidates = []
            for candidate_origin in allowed:
                trial_rows, trial_cycles = core.pack(candidate_origin, queues[candidate_origin] + [{**job, "manual_order": None}])
                current = next(row for row in trial_rows if row["key"] == job["key"])
                candidates.append(((0 if current["status"] == "CUMPLE" else 1, max(0, -current["slack_h"]), len(trial_cycles)), candidate_origin))
            origin = min(candidates)[1]
        counters[(origin, job["process_date"].date())] += 1
        order = int(plan.get("orden_manual") or counters[(origin, job["process_date"].date())])
        queues[origin].append({**job, "manual_order": order})
    rows = []
    for origin in ("NELLY", "FORTALEZA"):
        packed, _ = core.pack(origin, queues[origin])
        rows.extend(packed)
    for position, job in enumerate(direct, 1):
        dispatch = core.release_time(job)
        arrival = dispatch + timedelta(hours=job["transit_h"])
        slack = (job["appointment"] - arrival).total_seconds() / 3600
        rows.append({**job, "assigned_origin": "DIRECTO SIN PREENFRIO", "order": position,
                     "start": dispatch, "end": dispatch, "dispatch": dispatch, "arrival": arrival,
                     "status": "CUMPLE" if slack >= 0 else "NO CUMPLE"})
    return rows


def read_movements(path, year, week):
    frame = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    c = headers(frame)
    columns = {
        "sku": pick(c, "SKU", required=True), "lote": pick(c, "Lote", required=True),
        "farm": pick(c, "Nombre de Organización", "Finca", required=True),
        "destination": pick(c, "Destino", required=True), "pallets": pick(c, "Pallets", required=True),
        "process_date": pick(c, "Fecha de proceso", required=True), "week": pick(c, "SEMANA"),
        "appointment": pick(c, "FECHA CITA"), "origin": pick(c, "PREENFRIO"),
        "ingress_date": pick(c, "FECHA INGRESO AL PREENFRIO"),
        "ingress_time": pick(c, "HORA INGRESO AL PREENFRIO"), "cycle": pick(c, "HORAS DE CICLO"),
        "dispatch_date": pick(c, "FECHA DE DESPACHO REAL"), "dispatch_time": pick(c, "HORA DE DESPACHO REAL"),
        "status": pick(c, "CUMPLIMIENTO"), "observations": pick(c, "OBSERVACIONES"),
    }
    result = []
    for index, source in frame.iterrows():
        process_date = as_datetime(source[columns["process_date"]])
        if not process_date or process_date.isocalendar()[:2] != (year, week):
            continue
        try:
            pallets = float(source[columns["pallets"]])
        except (TypeError, ValueError):
            continue
        row = {
            "source_row": index + 2, "sku": source[columns["sku"]], "lote": source[columns["lote"]],
            "farm": source[columns["farm"]], "destination": source[columns["destination"]],
            "pallets": pallets, "process_date": process_date.replace(hour=0, minute=0, second=0, microsecond=0),
            "appointment": as_datetime(source[columns["appointment"]]) if columns["appointment"] else None,
            "origin": normalize_origin(source[columns["origin"]], source[columns["destination"]]) if columns["origin"] else None,
            "start": combine_date_time(source[columns["ingress_date"]], source[columns["ingress_time"]]) if columns["ingress_date"] and columns["ingress_time"] else None,
            "dispatch": combine_date_time(source[columns["dispatch_date"]], source[columns["dispatch_time"]]) if columns["dispatch_date"] and columns["dispatch_time"] else None,
            "status": str(source[columns["status"]]).strip() if columns["status"] and not pd.isna(source[columns["status"]]) else None,
            "observations": str(source[columns["observations"]]).strip() if columns["observations"] and not pd.isna(source[columns["observations"]]) else None,
        }
        duration = hours(source[columns["cycle"]]) if columns["cycle"] else 0
        row["end"] = row["start"] + timedelta(hours=duration) if row["start"] and duration > 0 else None
        row["base"] = base_key(row)
        result.append(row)
    occurrences = defaultdict(int)
    for row in result:
        occurrences[row["base"]] += 1
        row["occurrence"] = occurrences[row["base"]]
    return result


def match_real(planned, movements):
    groups = defaultdict(list)
    for movement in movements:
        groups[movement["base"]].append(movement)
    used = set()
    matches = {}
    for row in planned:
        candidates = [r for r in groups[base_key(row)] if r["source_row"] not in used]
        if not candidates:
            continue
        candidates.sort(key=lambda r: (
            0 if abs(float(r["pallets"]) - float(row["pallets"])) < 1e-6 else 1,
            abs(float(r["pallets"]) - float(row["pallets"])), r["source_row"]
        ))
        selected = candidates[0]
        used.add(selected["source_row"])
        matches[row["key"]] = selected
    return matches


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("planeacion")
    parser.add_argument("movimientos")
    parser.add_argument("--anio", type=int, default=2026)
    parser.add_argument("--semana", type=int, required=True)
    parser.add_argument("--aplicar", action="store_true")
    args = parser.parse_args()
    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY") or os.environ.get("SUPABASE_SECRET_KEY")
    if not url or not key:
        raise SystemExit("Faltan SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY")
    db = create_client(url, key)
    closure_rows = (db.table("cierres_semanales").select("*").eq("anio", args.anio).eq("semana", args.semana).eq("estado", "CERRADA").limit(1).execute().data or [])
    if not closure_rows:
        raise SystemExit("No existe un cierre CERRADA para la semana solicitada")
    closure = closure_rows[0]
    jobs = build_jobs(db, args.anio, args.semana)
    planned = calculate_week_plan(jobs)
    movements = read_movements(Path(args.movimientos), args.anio, args.semana)
    real_by_id = match_real(planned, movements)
    payloads = []
    for row in planned:
        real = real_by_id.get(row["key"])
        planned_origin = row.get("assigned_origin")
        planned_start, planned_end = row.get("start"), row.get("end")
        destination = row.get("destination")
        if real:
            real_origin = real.get("origin") or planned_origin
            shown_origin = real_origin
            shown_start = real.get("start") or planned_start
            shown_end = real.get("end") or planned_end
            shown_dispatch = real.get("dispatch") or row.get("dispatch") or planned_end
            source = "REAL" if real.get("start") and real.get("end") else "MIXTO"
            state = "TERMINADO" if real.get("start") and real.get("end") else "TERMINADO_SIN_TIEMPO_CAMARA"
        else:
            real_origin = None
            shown_origin, shown_start, shown_end = planned_origin, planned_start, planned_end
            shown_dispatch = row.get("dispatch") or planned_end
            source, state = "PLANEACION_SIN_REAL", "SIN_REGISTRO_REAL"
        if norm(destination) in {"CEDA", "JAPON", "JAPÓN"}:
            shown_origin = "DIRECTO SIN PREENFRIO"
            state = "DIRECTO_SIN_PREENFRIO"
        snapshot_key = f"{row['key']}|{args.anio}|{args.semana}"
        payloads.append({
            "cierre_id": closure["id"], "proceso_id": row["key"], "anio": args.anio, "semana": args.semana,
            "snapshot_key": snapshot_key, "fecha_proceso": iso(row["process_date"]), "fecha_cita": iso(row["appointment"]),
            "finca": row.get("farm"), "cliente": row.get("client"), "destino": destination,
            "sku": row.get("sku"), "lote": row.get("lote"), "pallets": row.get("pallets"),
            "orden_planeado": row.get("order"), "preenfrio_planeado": planned_origin,
            "ingreso_planeado": iso(planned_start), "fin_planeado": iso(planned_end),
            "despacho_planeado": iso(row.get("dispatch") or planned_end), "llegada_planeada": iso(row.get("arrival")),
            "cumplimiento_planeado": row.get("status"), "preenfrio_real": real.get("origin") if real else None,
            "ingreso_real": iso(real.get("start")) if real else None, "fin_real": iso(real.get("end")) if real else None,
            "despacho_real": iso(real.get("dispatch")) if real else None, "cumplimiento_real": real.get("status") if real else None,
            "estado_cierre": state, "tiene_real": bool(real), "fuente_tiempo": source,
            "observaciones": real.get("observations") if real else "Sin movimiento real; conserva la planeación"
        })
    summary = {
        "semana": args.semana, "planeados": len(planned), "con_real": sum(bool(real_by_id.get(r["key"])) for r in planned),
        "sin_real": sum(not bool(real_by_id.get(r["key"])) for r in planned), "modo": "APLICAR" if args.aplicar else "VISTA_PREVIA"
    }
    print(pd.Series(summary).to_json(force_ascii=False, indent=2))
    if not args.aplicar:
        print("Vista previa. Agrega --aplicar para guardar el snapshot.")
        return
    db.table("cierres_semanales_detalle").delete().eq("cierre_id", closure["id"]).execute()
    for start in range(0, len(payloads), 200):
        db.table("cierres_semanales_detalle").insert(payloads[start:start + 200]).execute()
    print("Snapshot histórico guardado. Planeación preservada y datos reales superpuestos.")


if __name__ == "__main__":
    main()
