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


def fast_automatic(jobs):
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
            rows, cycles = core.pack(origin, queues[origin] + [job])
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


def calculate(transient_overrides=None):
    """Calcula el plan y aplica inmediatamente los cambios enviados por las flechas.

    Los cambios tambien se guardan en Supabase, pero la respuesta no depende de
    volver a leer la base para reflejar el nuevo orden.
    """
    jobs, plans = load_jobs_and_plan()
    transient_overrides = transient_overrides or {}
    automatic_assignments = fast_automatic(jobs)
    queues = {"NELLY": [], "FORTALEZA": []}
    direct = []

    for job in jobs:
        if job["key"] not in automatic_assignments:
            continue

        plan = dict(plans.get(job["key"], {}))
        immediate_change = transient_overrides.get(job["key"])
        if immediate_change:
            immediate_origin = immediate_change.get("origin")
            immediate_order = immediate_change.get("order")
            if immediate_order is not None:
                plan["orden_manual"] = int(immediate_order)
            if immediate_origin == "DIRECTO SIN PREENFRIO":
                plan["preenfrio_manual"] = None
                plan["tipo_asignacion"] = "DIRECTO_SIN_PREENFRIO"
            elif immediate_origin in ("NELLY", "FORTALEZA"):
                plan["preenfrio_manual"] = immediate_origin
                plan["tipo_asignacion"] = "MANUAL"

        suggested = automatic_assignments[job["key"]]
        permitted = core.allowed(job)
        assignment_type = plan.get("tipo_asignacion") or "AUTOMATICA"

        if assignment_type == "DIRECTO_SIN_PREENFRIO":
            origin = "DIRECTO SIN PREENFRIO"
        else:
            if not permitted:
                continue
            origin = plan.get("preenfrio_manual") or suggested["origin"]
            if origin not in permitted:
                origin = suggested["origin"] if suggested["origin"] in permitted else permitted[0]

        try:
            order = int(plan.get("orden_manual") or suggested["order"])
        except (TypeError, ValueError):
            order = int(suggested["order"])

        item = {
            **job,
            "manual_order": order,
            "suggested_origin": suggested["origin"],
            "suggested_order": suggested["order"],
        }
        if origin == "DIRECTO SIN PREENFRIO":
            direct.append(item)
        else:
            queues[origin].append(item)

    rows, cycles = [], []
    for origin in ("NELLY", "FORTALEZA"):
        packed_rows, packed_cycles = core.pack(origin, queues[origin])
        rows.extend(packed_rows)
        cycles.extend(packed_cycles)

    for order, job in enumerate(
        sorted(direct, key=lambda item: (core.release_time(item), core.deadline(item))), 1
    ):
        dispatch = core.release_time(job)
        arrival = dispatch + timedelta(hours=job["transit_h"])
        slack = (job["appointment"] - arrival).total_seconds() / 3600
        rows.append(
            {
                **job,
                "assigned_origin": "DIRECTO SIN PREENFRIO",
                "assigned_distance_km": None,
                "order": order,
                "start": dispatch,
                "end": dispatch,
                "pre_cooling_end": dispatch,
                "dispatch": dispatch,
                "arrival": arrival,
                "conservation_wait_h": 0.0,
                "portions": [],
                "slack_h": slack,
                "status": "CUMPLE" if slack >= 0 else "NO CUMPLE",
            }
        )
    return rows, cycles, plans


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


def serialise(rows, cycles):
    row_fields = [
        "key", "farm", "destination", "client", "process_date", "appointment",
        "transit_h", "critical", "pallets", "week", "assigned_origin",
        "assigned_distance_km", "distance_options_km", "location_found", "order",
        "start", "end", "pre_cooling_end", "dispatch", "arrival",
        "conservation_wait_h", "portions", "slack_h", "status",
    ]
    cycle_fields = [
        "cycle_no", "origin", "process_date", "start", "end", "duration_h",
        "capacity", "used_pallets", "free_pallets", "loads",
    ]
    return {
        "ok": True,
        "source": "SUPABASE",
        "version": "2.5.3",
        "data": [
            {field: json_value(row.get(field)) for field in row_fields}
            for row in rows
        ],
        "cycles": [
            {field: json_value(cycle.get(field)) for field in cycle_fields}
            for cycle in cycles
        ],
    }


def payload(force=False, transient_overrides=None):
    with LOCK:
        now = clock.monotonic()
        if (
            not force
            and CACHE["payload"] is not None
            and now - CACHE["time"] < CACHE_SECONDS
        ):
            return CACHE["payload"]
        rows, cycles, _ = calculate(transient_overrides)
        CACHE.update(payload=serialise(rows, cycles), time=now)
        return CACHE["payload"]


def save_overrides(overrides):
    modified_at = datetime.now().astimezone().isoformat()
    for process_id, change in overrides.items():
        origin = change.get("origin")
        values = {
            "orden_manual": int(change.get("order")) if change.get("order") else None,
            "modificado_por": "DASHBOARD",
            "modificado_en": modified_at,
        }
        if origin == "DIRECTO SIN PREENFRIO":
            values.update(preenfrio_manual=None, tipo_asignacion="DIRECTO_SIN_PREENFRIO")
        elif origin in ("NELLY", "FORTALEZA"):
            values.update(preenfrio_manual=origin, tipo_asignacion="MANUAL")
        DB.table("procesos_planificacion").update(values).eq("proceso_id", process_id).execute()


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
    print(f"Planeador Supabase V2.5.3: http://{HOST}:{PORT}")
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
