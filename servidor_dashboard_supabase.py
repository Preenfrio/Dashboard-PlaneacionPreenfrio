#!/usr/bin/env python3
"""Servidor del Planeador usando Supabase como fuente operativa.

Reutiliza el optimizador de servidor_dashboard_v2.py y conserva el HTML actual.
Excel se usa solo para CATALOGO y UBICACIONES al arrancar; los procesos y ajustes
manuales se leen/escriben en Supabase.
"""
from __future__ import annotations

import io
import hmac
import base64
import json
import os
import sys
import subprocess
import threading
import tempfile
import time as clock
from datetime import datetime, timedelta, time
from http.server import ThreadingHTTPServer
from urllib.parse import urlparse
from pathlib import Path

import pandas as pd
from supabase import create_client

import servidor_dashboard_v2 as core

HOST = "0.0.0.0" if os.environ.get("PORT") else "127.0.0.1"
PORT = int(os.environ.get("PORT", "8010"))
CACHE_SECONDS = int(os.environ.get("CACHE_SECONDS", "45"))
SYNC_TOKEN = os.environ.get("POWER_AUTOMATE_SYNC_TOKEN", "")
MAX_RETIRED_RATIO = float(os.environ.get("MAX_RETIRED_RATIO", "0.20"))
IMPORTER = Path(__file__).with_name("importador_incremental_supabase.py")

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
    # procesos_fuente y procesos_planificacion comparten proceso_id, pero no tienen
    # una FK directa entre sí. Se consultan por separado y se unen en Python.
    source = fetch_all(
        "procesos_fuente",
        "*,procesos!inner(id,activo_fuente)",
    )
    plan_rows = fetch_all("procesos_planificacion", "*")
    plans = {row["proceso_id"]: row for row in plan_rows}
    jobs = []

    for index, row in enumerate(source, 1):
        process_meta = row.pop("procesos")
        if not process_meta.get("activo_fuente", True):
            continue

        destination = str(row.get("destino") or "").strip()
        info = CATALOGUE.get(core.norm(destination))
        process_date = parse_dt(row.get("fecha_proceso"))
        appointment = parse_dt(row.get("fecha_cita"))
        pallets = float(row.get("pallets") or 0)
        if not info or not process_date or not appointment or pallets <= 0:
            continue

        farm = str(row.get("finca") or "").strip()
        farm_coord = LOCATIONS.get(core.norm(farm))
        distances = {
            origin: core.haversine_km(farm_coord, coord)
            for origin, coord in PRECOOLERS.items()
        }
        process_id = process_meta["id"]
        jobs.append({
            "key": process_id,
            "source": index,
            "farm": farm,
            "destination": destination,
            "client": core.norm(row.get("cliente") or info["client"]),
            "process_date": process_date.replace(hour=0, minute=0, second=0, microsecond=0),
            "appointment": appointment,
            "transit_h": float(row.get("horas_transito") if row.get("horas_transito") is not None else info["transit_h"]),
            "critical": int(row.get("nivel_critico") if row.get("nivel_critico") is not None else info["critical"]),
            "pallets": pallets,
            "week": str(row.get("semana") or process_date.isocalendar().week),
            "cycle_options": info["cycle"],
            "distance_options_km": distances,
            "location_found": farm_coord is not None,
            "ov": row.get("ov"),
            "sku": row.get("sku"),
            "lote": row.get("lote"),
        })
    return jobs, plans

def fast_automatic(jobs):
    """Asignador rápido para volumen alto.

    Mantiene la prioridad por disponibilidad, criticidad y cita. Para cada proceso
    prueba solo las cámaras permitidas, no todas las posiciones posibles de toda
    la cola, evitando el crecimiento combinatorio del optimizador exhaustivo.
    """
    queues = {"NELLY": [], "FORTALEZA": []}
    assignments = {}
    ordered = sorted(jobs, key=lambda j: (
        core.release_time(j), j.get("critical", 99), j["appointment"], j["source"]
    ))
    for job in ordered:
        best = None
        for origin in core.allowed(job):
            trial = list(queues[origin]) + [job]
            trial.sort(key=lambda j: (
                core.release_time(j), j.get("critical", 99), j["appointment"], j["source"]
            ))
            rows, cycles = core.pack(origin, trial)
            candidate = next(row for row in rows if row["key"] == job["key"])
            protected_late = sum(
                row["status"] == "NO CUMPLE" and row.get("critical", 99) <= 4
                for row in rows
            )
            protected_delay = sum(
                max(0.0, -row["slack_h"])
                for row in rows if row.get("critical", 99) <= 4
            )
            distance = job["distance_options_km"].get(origin)
            score = (
                protected_late,
                round(protected_delay, 6),
                0 if candidate["status"] == "CUMPLE" else 1,
                max(0.0, -candidate["slack_h"]),
                distance if distance is not None else 1000000.0,
                len(cycles),
            )
            if best is None or score < best[0]:
                best = (score, origin)
        if best is None:
            continue
        origin = best[1]
        queues[origin].append(job)
        same_day = [x for x in queues[origin] if x["process_date"].date() == job["process_date"].date()]
        assignments[job["key"]] = {"origin": origin, "order": len(same_day)}
    return assignments


def calculate():
    jobs, plans = load_jobs_and_plan()
    automatic = fast_automatic(jobs)
    queues = {"NELLY": [], "FORTALEZA": []}
    direct = []

    for job in jobs:
        if job["key"] not in automatic:
            continue
        plan = plans.get(job["key"], {})
        suggested = automatic[job["key"]]
        manual_origin = plan.get("preenfrio_manual")
        kind = plan.get("tipo_asignacion") or "AUTOMATICA"
        if kind == "DIRECTO_SIN_PREENFRIO":
            origin = "DIRECTO SIN PREENFRIO"
        else:
            origin = manual_origin or suggested["origin"]
        order = int(plan.get("orden_manual") or suggested["order"])
        item = {**job, "manual_order": order, "suggested_origin": suggested["origin"], "suggested_order": suggested["order"]}
        if origin == "DIRECTO SIN PREENFRIO":
            direct.append(item)
            continue
        if job["client"] == "SMART":
            origin = "NELLY"
        if origin not in core.allowed(job):
            origin = suggested["origin"]
        queues[origin].append(item)

    rows, cycles = [], []
    for origin in ("NELLY", "FORTALEZA"):
        queues[origin].sort(key=lambda j: (core.release_time(j), j.get("manual_order", 9999), j["critical"], j["appointment"], j["source"]))
        packed_rows, packed_cycles = core.pack(origin, queues[origin])
        rows.extend(packed_rows)
        cycles.extend(packed_cycles)

    for order, job in enumerate(sorted(direct, key=lambda j: (core.release_time(j), j.get("manual_order", 9999))), 1):
        dispatch = core.release_time(job)
        arrival = dispatch + timedelta(hours=job["transit_h"])
        slack = (job["appointment"] - arrival).total_seconds() / 3600
        rows.append({**job, "assigned_origin": "DIRECTO SIN PREENFRIO", "assigned_distance_km": None,
                     "order": order, "start": dispatch, "end": dispatch, "arrival": arrival,
                     "slack_h": slack, "status": "CUMPLE" if slack >= 0 else "NO CUMPLE"})
    return rows, cycles, plans


def persist_results(rows, cycles):
    planning = []
    for row in rows:
        planning.append({
            "proceso_id": row["key"],
            "preenfrio_sugerido": row.get("suggested_origin") or (row["assigned_origin"] if row["assigned_origin"] != "DIRECTO SIN PREENFRIO" else None),
            "orden_sugerido": row.get("suggested_order") or row["order"],
            "fecha_ingreso": row["start"].isoformat(timespec="seconds"),
            "fecha_despacho": row["end"].isoformat(timespec="seconds"),
            "llegada_estimada": row["arrival"].isoformat(timespec="seconds"),
            "cumplimiento": row["status"],
            "holgura_horas": round(row["slack_h"], 6),
            "distancia_km": None if row.get("assigned_distance_km") is None else round(row["assigned_distance_km"], 6),
        })
    # UPDATE, no UPSERT: evita que columnas manuales ausentes se conviertan en NULL.
    for item in planning:
        process_id = item.pop("proceso_id")
        DB.table("procesos_planificacion").update(item).eq("proceso_id", process_id).execute()

    # Los ciclos son resultados derivados; se regeneran completos para evitar residuos.
    DB.table("ciclos_procesos").delete().neq("ciclo_id", "00000000-0000-0000-0000-000000000000").execute()
    DB.table("ciclos_preenfrio").delete().neq("id", "00000000-0000-0000-0000-000000000000").execute()
    for cycle in cycles:
        inserted = DB.table("ciclos_preenfrio").insert({
            "preenfrio": cycle["origin"], "numero_ciclo": cycle["cycle_no"],
            "inicio": cycle["start"].isoformat(timespec="seconds"), "fin": cycle["end"].isoformat(timespec="seconds"),
            "capacidad": cycle["capacity"], "tarimas_usadas": cycle["used_pallets"], "tarimas_libres": cycle["free_pallets"],
        }).execute().data[0]
        links = [{"ciclo_id": inserted["id"], "proceso_id": load["key"], "tarimas": load["pallets"]} for load in cycle["loads"]]
        if links:
            DB.table("ciclos_procesos").insert(links).execute()


def serialise(rows, cycles):
    def iso(value):
        return value.isoformat(timespec="minutes") if isinstance(value, datetime) else value
    row_fields = ["key", "farm", "destination", "client", "process_date", "appointment", "transit_h", "critical", "pallets", "week", "assigned_origin", "assigned_distance_km", "distance_options_km", "location_found", "order", "start", "end", "arrival", "slack_h", "status"]
    cycle_fields = ["cycle_no", "origin", "process_date", "start", "end", "duration_h", "capacity", "used_pallets", "free_pallets", "loads"]
    return {"ok": True, "source": "SUPABASE", "data": [{f: iso(row[f]) for f in row_fields} for row in rows], "cycles": [{f: iso(cycle[f]) for f in cycle_fields} for cycle in cycles]}


def payload(force=False):
    with LOCK:
        now = clock.monotonic()
        if not force and CACHE["payload"] is not None and now - CACHE["time"] < CACHE_SECONDS:
            return CACHE["payload"]
        rows, cycles, _ = calculate()
        # La lectura del dashboard no escribe cientos de filas ni regenera tablas.
        # Las decisiones manuales se guardan en save_overrides().
        CACHE.update(payload=serialise(rows, cycles), time=now)
        return CACHE["payload"]


def save_overrides(overrides):
    now = datetime.now().astimezone().isoformat()
    for process_id, change in overrides.items():
        origin = change.get("origin")
        order = change.get("order")
        values = {"proceso_id": process_id, "orden_manual": int(order) if order else None,
                  "modificado_por": "DASHBOARD", "modificado_en": now}
        if origin == "DIRECTO SIN PREENFRIO":
            values.update(preenfrio_manual=None, tipo_asignacion="DIRECTO_SIN_PREENFRIO")
        elif origin in ("NELLY", "FORTALEZA"):
            values.update(preenfrio_manual=origin, tipo_asignacion="MANUAL")
        # La fila ya existe desde la importación inicial. UPDATE protege los demás campos.
        process_id_value = values.pop("proceso_id")
        DB.table("procesos_planificacion").update(values).eq("proceso_id", process_id_value).execute()


def export_bytes():
    p = payload()
    frame = pd.DataFrame([{
        "CLAVE": x["key"], "FECHA_PROCESO": x["process_date"][:10], "CLIENTE": x["client"],
        "FINCA": x["farm"], "DESTINO": x["destination"], "TARIMAS": x["pallets"],
        "PREENFRIO_ASIGNADO": x["assigned_origin"], "DISTANCIA_KM": x["assigned_distance_km"],
        "ORDEN": x["order"], "FECHA_CITA": x["appointment"], "FECHA_INGRESO": x["start"],
        "FECHA_DESPACHO": x["end"], "ESTADO": x["status"],
    } for x in p["data"]])
    buffer = io.BytesIO()
    frame.to_excel(buffer, sheet_name="PLAN_AJUSTADO", index=False, engine="openpyxl")
    return buffer.getvalue()



def _summary_from_output(output):
    start = output.find("{")
    if start < 0:
        raise ValueError("El importador no devolvió un resumen JSON")
    summary, _ = json.JSONDecoder().raw_decode(output[start:])
    return summary


def _run_incremental(excel_path, apply=False):
    command = [sys.executable, str(IMPORTER), str(excel_path)]
    if apply:
        command.append("--aplicar")
    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=300,
        env=os.environ.copy(),
    )
    output = (completed.stdout or "") + "\n" + (completed.stderr or "")
    if completed.returncode != 0:
        raise RuntimeError(output.strip() or "El importador terminó con error")
    return _summary_from_output(output)


def synchronise_from_power_automate(body, supplied_token):
    if not SYNC_TOKEN:
        return {"ok": False, "aplicado": False, "motivo": "TOKEN_NO_CONFIGURADO"}, 503
    if not hmac.compare_digest(str(supplied_token or ""), SYNC_TOKEN):
        return {"ok": False, "aplicado": False, "motivo": "TOKEN_INVALIDO"}, 401
    if not IMPORTER.is_file():
        return {"ok": False, "aplicado": False, "motivo": "IMPORTADOR_NO_ENCONTRADO"}, 500

    file_name = str(body.get("nombre_archivo") or "PLANEACION_SEMANAL 1.xlsx")
    encoded = body.get("contenido_base64")
    if not encoded:
        return {"ok": False, "aplicado": False, "motivo": "CONTENIDO_VACIO"}, 400
    try:
        binary = base64.b64decode(encoded, validate=True)
    except Exception:
        return {"ok": False, "aplicado": False, "motivo": "BASE64_INVALIDO"}, 400
    if not binary.startswith(b"PK"):
        return {"ok": False, "aplicado": False, "motivo": "XLSX_INVALIDO"}, 400

    with tempfile.TemporaryDirectory(prefix="planeacion_sync_") as folder:
        excel_path = Path(folder) / "PLANEACION_SEMANAL 1.xlsx"
        excel_path.write_bytes(binary)
        preview = _run_incremental(excel_path, apply=False)

        valid = int(preview.get("filas_validas", 0))
        new = int(preview.get("nuevos", 0))
        modified = int(preview.get("modificados", 0))
        unchanged = int(preview.get("sin_cambios", 0))
        retired = int(preview.get("retirados", 0))
        ambiguous = int(preview.get("ambiguos", 0))
        total_existing = modified + unchanged + retired
        retired_ratio = retired / total_existing if total_existing else 0.0

        response = {
            "ok": True,
            "aplicado": False,
            "motivo": None,
            "archivo": file_name,
            "filas_validas": valid,
            "filas_rechazadas": int(preview.get("filas_rechazadas", 0)),
            "nuevos": new,
            "modificados": modified,
            "sin_cambios": unchanged,
            "retirados": retired,
            "ambiguos": ambiguous,
            "fecha_sincronizacion": datetime.now().astimezone().isoformat(),
        }
        if valid <= 0:
            response.update(ok=False, motivo="SIN_FILAS_VALIDAS")
            return response, 422
        if ambiguous > 0:
            response["motivo"] = "CONCILIACIONES_AMBIGUAS"
            return response, 409
        if retired_ratio > MAX_RETIRED_RATIO:
            response["motivo"] = "RETIRADOS_SUPERAN_LIMITE"
            response["proporcion_retirados"] = round(retired_ratio, 6)
            return response, 409
        if new + modified + retired == 0:
            response["motivo"] = "SIN_CAMBIOS"
            return response, 200

        _run_incremental(excel_path, apply=True)
        response["aplicado"] = True
        response["motivo"] = "APLICADO"
        response["fecha_sincronizacion"] = datetime.now().astimezone().isoformat()
        with LOCK:
            CACHE.update(payload=None, time=0.0)
        return response, 200


class Handler(core.H):
    def do_GET(self):
        path = urlparse(self.path).path
        try:
            if path == "/":
                return self.sendx(core.HTML.read_bytes(), "text/html; charset=utf-8")
            if path == "/api/datos":
                return self.sendx(json.dumps(payload(), ensure_ascii=False).encode(), "application/json; charset=utf-8")
            if path == "/api/exportar":
                return self.sendx(export_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", extra={"Content-Disposition": 'attachment; filename="PLAN_AJUSTADO.xlsx"'})
        except Exception as exc:
            return self.sendx(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False).encode(), "application/json; charset=utf-8", 500)
        return self.sendx(b"No encontrado", code=404)

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            raw = self.rfile.read(int(self.headers.get("Content-Length", "0")))
            if path == "/api/sincronizar-planeacion":
                body = json.loads(raw or b"{}")
                result, status = synchronise_from_power_automate(
                    body, self.headers.get("X-Sync-Token")
                )
                return self.sendx(
                    json.dumps(result, ensure_ascii=False).encode(),
                    "application/json; charset=utf-8",
                    status,
                )
            if path == "/api/replan":
                body = json.loads(raw or b"{}")
                save_overrides(body.get("overrides", {}))
                result = payload(force=True)
                return self.sendx(json.dumps(result, ensure_ascii=False).encode(), "application/json; charset=utf-8")
            if path == "/api/importar":
                return self.sendx(json.dumps({"ok": False, "error": "La importación base ahora se realiza con importador_incremental_supabase.py"}, ensure_ascii=False).encode(), "application/json; charset=utf-8", 409)
        except Exception as exc:
            return self.sendx(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False).encode(), "application/json; charset=utf-8", 500)
        return self.sendx(b"No encontrado", code=404)


if __name__ == "__main__":
    print(f"Planeador Supabase: http://{HOST}:{PORT}")
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
