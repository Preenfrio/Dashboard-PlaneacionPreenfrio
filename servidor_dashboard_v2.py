#!/usr/bin/env python3
from __future__ import annotations

import io
import json
import math
import os
import re
import threading
import webbrowser
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


def norm(value):
    return " ".join(str(value if value is not None else "").strip().upper().split())


def col(df, name, required=True):
    wanted = norm(name)
    for column in df.columns:
        if norm(column) == wanted:
            return column
    if required:
        raise ValueError(f"Falta columna: {name}")
    return None


def as_dt(value):
    parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
    return None if pd.isna(parsed) else parsed.to_pydatetime()


def hours(value):
    if pd.isna(value) or value == "":
        return 0.0
    if isinstance(value, (datetime, time)):
        return value.hour + value.minute / 60 + value.second / 3600
    if isinstance(value, timedelta):
        return value.total_seconds() / 3600
    if isinstance(value, (int, float)):
        number = float(value)
        return number * 24 if 0 <= number < 1.5 else number
    parts = str(value).strip().split(":")
    try:
        if len(parts) > 1:
            return float(parts[0]) + float(parts[1]) / 60 + (float(parts[2]) / 3600 if len(parts) > 2 else 0)
        return float(parts[0])
    except (TypeError, ValueError):
        return 0.0


def parse_coordinate(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    text = str(value).strip().upper().replace("–", "-").replace("—", "-")
    if not text:
        return None
    dms = re.findall(r'(\d{1,3})\s*°\s*(\d{1,2})\s*[\'’]\s*(\d+(?:\.\d+)?)\s*[\"”]?\s*([NSEW])', text)
    if len(dms) >= 2:
        values = []
        for degree, minute, second, hemisphere in dms[:2]:
            number = float(degree) + float(minute) / 60 + float(second) / 3600
            if hemisphere in ("S", "W"):
                number *= -1
            values.append(number)
        return tuple(values)
    numbers = re.findall(r"[-+]?\d{1,3}(?:\.\d+)", text)
    if len(numbers) >= 2:
        latitude, longitude = float(numbers[0]), float(numbers[1])
        if "W" in text and longitude > 0:
            longitude *= -1
        if "S" in text and latitude > 0:
            latitude *= -1
        if -90 <= latitude <= 90 and -180 <= longitude <= 180:
            return latitude, longitude
    return None


def haversine_km(point_a, point_b):
    if not point_a or not point_b:
        return None
    lat1, lon1 = map(math.radians, point_a)
    lat2, lon2 = map(math.radians, point_b)
    dlat, dlon = lat2 - lat1, lon2 - lon1
    h = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return 6371.0088 * 2 * math.asin(math.sqrt(h))


def read_locations():
    frame = pd.read_excel(EXCEL, sheet_name="UBICACIONES", engine="openpyxl")
    name_column = col(frame, "Trazabilidad de finca")
    coordinate_column = col(frame, "Coordenadas")
    locations = {}
    for _, row in frame.iterrows():
        name = norm(row[name_column])
        coordinate = parse_coordinate(row[coordinate_column])
        if name and coordinate:
            locations[name] = coordinate
    precoolers = {
        "NELLY": locations.get("PREENFRIO NELLY"),
        "FORTALEZA": locations.get("PREENFRIO FORTALEZA"),
    }
    return locations, precoolers


def read_catalogue():
    frame = pd.read_excel(EXCEL, sheet_name="CATALOGO", engine="openpyxl")
    columns = {name: col(frame, name) for name in [
        "CLIENTE", "DESTINO", "TRÁNSITO (HRS)",
        "CICLO HORAS FORTALEZA", "CICLO HORAS NELLY", "CRITICIDAD",
    ]}
    result = {}
    for _, row in frame.iterrows():
        destination = norm(row[columns["DESTINO"]])
        if not destination:
            continue
        try:
            critical = int(float(row[columns["CRITICIDAD"]]))
        except (TypeError, ValueError):
            critical = 99
        result[destination] = {
            "client": norm(row[columns["CLIENTE"]]),
            "transit_h": hours(row[columns["TRÁNSITO (HRS)"]]),
            "critical": critical,
            "cycle": {
                "NELLY": hours(row[columns["CICLO HORAS NELLY"]]),
                "FORTALEZA": hours(row[columns["CICLO HORAS FORTALEZA"]]),
            },
        }
    return result


def read_jobs():
    catalogue = read_catalogue()
    locations, precoolers = read_locations()
    frame = pd.read_excel(EXCEL, sheet_name="PLANEACION", engine="openpyxl")
    columns = {name: col(frame, name) for name in ["Finca", "Destino", "Pallets", "Fecha de proceso", "FECHA CITA"]}
    columns["SEMANA"] = col(frame, "SEMANA", False)
    columns["HORAS TRANSITO"] = col(frame, "HORAS TRANSITO", False)
    columns["NIVEL CRITICO"] = col(frame, "NIVEL CRITICO", False)
    jobs = {}
    for index, row in frame.iterrows():
        farm = str(row[columns["Finca"]]).strip()
        destination = str(row[columns["Destino"]]).strip()
        info = catalogue.get(norm(destination))
        process_date = as_dt(row[columns["Fecha de proceso"]])
        appointment = as_dt(row[columns["FECHA CITA"]])
        try:
            pallets = float(row[columns["Pallets"]])
        except (TypeError, ValueError):
            pallets = 0
        if not info or not process_date or not appointment or pallets <= 0:
            continue
        transit_h = hours(row[columns["HORAS TRANSITO"]]) if columns["HORAS TRANSITO"] else info["transit_h"]
        try:
            critical = int(float(row[columns["NIVEL CRITICO"]])) if columns["NIVEL CRITICO"] else info["critical"]
        except (TypeError, ValueError):
            critical = info["critical"]
        try:
            week = str(int(float(row[columns["SEMANA"]]))) if columns["SEMANA"] else str(process_date.isocalendar().week)
        except (TypeError, ValueError):
            week = str(process_date.isocalendar().week)
        farm_coordinate = locations.get(norm(farm))
        distances = {origin: haversine_km(farm_coordinate, coordinate) for origin, coordinate in precoolers.items()}
        key = f"{norm(farm)}|{norm(destination)}|{process_date.date()}|{appointment.date()}|{info['client']}"
        if key not in jobs:
            jobs[key] = {
                "key": key, "source": index + 2, "farm": farm, "destination": destination,
                "client": info["client"], "process_date": process_date.replace(hour=0, minute=0, second=0, microsecond=0),
                "appointment": appointment, "transit_h": transit_h, "critical": critical,
                "pallets": pallets, "week": week, "cycle_options": info["cycle"],
                "distance_options_km": distances, "location_found": farm_coordinate is not None,
            }
        else:
            jobs[key]["pallets"] += pallets
    return list(jobs.values())


def allowed(job):
    """Cámaras permitidas, respetando la separación física Nelly/Fortaleza."""
    farm = norm(job.get("farm", ""))
    client = norm(job.get("client", ""))
    cycle_options = job.get("cycle_options") or {}
    nelly_enabled = hours(cycle_options.get("NELLY", 0)) > 0
    fortaleza_enabled = hours(cycle_options.get("FORTALEZA", 0)) > 0

    is_nelly_farm = farm.startswith("A01 003") or "DOÑA NELLY" in farm or "DONA NELLY" in farm
    if is_nelly_farm:
        return ["NELLY"] if nelly_enabled else []

    is_fortaleza_farm = farm.startswith("A01 002") or farm == "FORTALEZA" or "A01 002 FORTALEZA" in farm
    if is_fortaleza_farm:
        return ["FORTALEZA"] if fortaleza_enabled else []

    if client == "SMART":
        return ["NELLY"] if nelly_enabled else []

    result = []
    if nelly_enabled:
        result.append("NELLY")
    if fortaleza_enabled:
        result.append("FORTALEZA")
    return result


def release_time(job):
    return datetime.combine(job["process_date"].date(), time(START_HOUR))


def pack(origin, sequence):
    """Construye una cola cronológica independiente para una sola cámara."""
    if not sequence:
        return [], []
    capacity = CAP[origin]
    sequence = sorted(sequence, key=lambda job: (
        release_time(job), job.get("manual_order", 9999), job.get("critical", 99), job["appointment"], job["source"]
    ))
    remaining = {job["key"]: job["pallets"] for job in sequence}
    first, last, cycles = {}, {}, []
    pending = list(sequence)
    current = min(release_time(job) for job in pending)
    while pending:
        available = [job for job in pending if release_time(job) <= current]
        if not available:
            current = min(release_time(job) for job in pending)
            available = [job for job in pending if release_time(job) <= current]
        available.sort(key=lambda job: (
            job.get("manual_order", 9999), job.get("critical", 99), job["appointment"], job["source"]
        ))
        space, loads, duration = capacity, [], 0.0
        for job in available:
            if space <= 1e-9:
                break
            quantity = min(space, remaining[job["key"]])
            if quantity <= 1e-9:
                continue
            loads.append({"key": job["key"], "destination": job["destination"], "farm": job.get("farm", ""), "pallets": quantity})
            duration = max(duration, hours(job["cycle_options"].get(origin, 0)))
            first.setdefault(job["key"], current)
            remaining[job["key"]] -= quantity
            space -= quantity
        if not loads:
            raise ValueError(f"No se pudo formar un ciclo válido para {origin}")
        end = current + timedelta(hours=duration)
        cycles.append({
            "cycle_no": len(cycles) + 1, "origin": origin,
            "process_date": current.replace(hour=0, minute=0, second=0, microsecond=0),
            "start": current, "end": end, "duration_h": duration, "capacity": capacity,
            "used_pallets": capacity - space, "free_pallets": space, "loads": loads,
        })
        completed = []
        for load in loads:
            key = load["key"]
            if remaining[key] <= 1e-9:
                last[key] = end
                completed.append(key)
        pending = [job for job in pending if job["key"] not in completed]
        current = end + timedelta(minutes=GAP_MIN)

    rows, counters = [], {}
    for job in sequence:
        end = last[job["key"]]
        arrival = end + timedelta(hours=job["transit_h"])
        slack = (job["appointment"] - arrival).total_seconds() / 3600
        day_key = (job["process_date"].date(), origin)
        counters[day_key] = counters.get(day_key, 0) + 1
        rows.append({
            **job, "assigned_origin": origin,
            "assigned_distance_km": job["distance_options_km"].get(origin),
            "order": counters[day_key], "start": first[job["key"]], "end": end,
            "arrival": arrival, "slack_h": slack,
            "status": "CUMPLE" if slack >= 0 else "NO CUMPLE",
        })
    return rows, cycles


def automatic(jobs):
    queues = {"NELLY": [], "FORTALEZA": []}
    assignments = {}
    ordered = sorted(jobs, key=lambda job: (release_time(job), job.get("critical", 99), job["appointment"], job["source"]))
    for job in ordered:
        best = None
        for origin in allowed(job):
            trial = list(queues[origin]) + [job]
            rows, cycles = pack(origin, trial)
            candidate = next(row for row in rows if row["key"] == job["key"])
            score = (
                0 if candidate["status"] == "CUMPLE" else 1,
                max(0.0, -candidate["slack_h"]),
                job["distance_options_km"].get(origin) if job["distance_options_km"].get(origin) is not None else 1000000.0,
                len(cycles),
            )
            if best is None or score < best[0]:
                best = (score, origin)
        if best is None:
            continue
        origin = best[1]
        queues[origin].append(job)
        same_day = [item for item in queues[origin] if item["process_date"].date() == job["process_date"].date()]
        assignments[job["key"]] = {"origin": origin, "order": len(same_day)}
    return assignments


def recalc(overrides=None):
    jobs = read_jobs()
    auto = automatic(jobs)
    selected_overrides = overrides if overrides is not None else STATE["overrides"]
    queues = {"NELLY": [], "FORTALEZA": []}
    direct = []
    for job in jobs:
        if job["key"] not in auto:
            continue
        chosen = selected_overrides.get(job["key"], auto[job["key"]])
        origin = chosen.get("origin") or auto[job["key"]]["origin"]
        manual_order = int(chosen.get("order") or auto[job["key"]]["order"])
        job_copy = {**job, "manual_order": manual_order}
        if origin == "DIRECTO SIN PREENFRIO":
            direct.append(job_copy)
            continue
        permitted = allowed(job)
        if origin not in permitted:
            origin = auto[job["key"]]["origin"] if auto[job["key"]]["origin"] in permitted else permitted[0]
        queues[origin].append(job_copy)
    rows, cycles = [], []
    for origin in ("NELLY", "FORTALEZA"):
        packed_rows, packed_cycles = pack(origin, queues[origin])
        rows.extend(packed_rows)
        cycles.extend(packed_cycles)
    for order, job in enumerate(sorted(direct, key=lambda item: (release_time(item), item.get("manual_order", 9999))), 1):
        dispatch = release_time(job)
        arrival = dispatch + timedelta(hours=job["transit_h"])
        slack = (job["appointment"] - arrival).total_seconds() / 3600
        rows.append({**job, "assigned_origin": "DIRECTO SIN PREENFRIO", "assigned_distance_km": None,
                     "order": order, "start": dispatch, "end": dispatch, "arrival": arrival,
                     "slack_h": slack, "status": "CUMPLE" if slack >= 0 else "NO CUMPLE"})
    return rows, cycles


def serial(rows, cycles):
    def iso(value):
        return value.isoformat(timespec="minutes") if isinstance(value, datetime) else value
    row_fields = ["key", "farm", "destination", "client", "process_date", "appointment", "transit_h", "critical", "pallets", "week", "assigned_origin", "assigned_distance_km", "distance_options_km", "location_found", "order", "start", "end", "arrival", "slack_h", "status"]
    cycle_fields = ["cycle_no", "origin", "process_date", "start", "end", "duration_h", "capacity", "used_pallets", "free_pallets", "loads"]
    return {"ok": True, "version": "2.4", "excel": EXCEL.name,
            "modified": datetime.fromtimestamp(EXCEL.stat().st_mtime).isoformat(timespec="seconds"),
            "data": [{field: iso(row[field]) for field in row_fields} for row in rows],
            "cycles": [{field: iso(cycle[field]) for field in cycle_fields} for cycle in cycles]}


def payload():
    mtime = EXCEL.stat().st_mtime
    if STATE["payload"] is None or STATE["mtime"] != mtime:
        STATE["mtime"] = mtime
        STATE["overrides"] = {}
        STATE["payload"] = serial(*recalc())
    return STATE["payload"]


def export_bytes():
    current = payload()
    frame = pd.DataFrame([{
        "CLAVE": item["key"], "FECHA_PROCESO": item["process_date"][:10], "CLIENTE": item["client"],
        "FINCA": item["farm"], "DESTINO": item["destination"], "TARIMAS": item["pallets"],
        "PREENFRIO_ASIGNADO": item["assigned_origin"], "DISTANCIA_KM": item["assigned_distance_km"],
        "ORDEN": item["order"], "FECHA_CITA": item["appointment"][:10], "FECHA_INGRESO": item["start"],
        "FECHA_DESPACHO": item["end"], "ESTADO": item["status"],
    } for item in current["data"]])
    buffer = io.BytesIO()
    frame.to_excel(buffer, sheet_name="PLAN_AJUSTADO", index=False, engine="openpyxl")
    return buffer.getvalue()


def import_overrides(raw):
    frame = pd.read_excel(io.BytesIO(raw), sheet_name="PLAN_AJUSTADO", engine="openpyxl")
    columns = {name: col(frame, name) for name in ["CLAVE", "PREENFRIO_ASIGNADO", "ORDEN"]}
    overrides = {}
    for _, row in frame.iterrows():
        overrides[str(row[columns["CLAVE"]]).strip()] = {
            "origin": norm(row[columns["PREENFRIO_ASIGNADO"]]),
            "order": int(float(row[columns["ORDEN"]])),
        }
    STATE["overrides"] = overrides
    STATE["payload"] = serial(*recalc(overrides))
    return STATE["payload"]


class H(BaseHTTPRequestHandler):
    def sendx(self, body, content_type="text/plain; charset=utf-8", code=200, extra=None):
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        for key, value in (extra or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/":
            return self.sendx(HTML.read_bytes(), "text/html; charset=utf-8")
        if path == "/api/datos":
            try:
                return self.sendx(json.dumps(payload(), ensure_ascii=False).encode(), "application/json; charset=utf-8")
            except Exception as exc:
                return self.sendx(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False).encode(), "application/json; charset=utf-8", 500)
        if path == "/api/exportar":
            return self.sendx(export_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", extra={"Content-Disposition": 'attachment; filename="PLAN_AJUSTADO_V2.xlsx"'})
        return self.sendx(b"No encontrado", code=404)

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            raw = self.rfile.read(int(self.headers.get("Content-Length", "0")))
            if path == "/api/replan":
                body = json.loads(raw or b"{}")
                STATE["overrides"] = body.get("overrides", {})
                STATE["payload"] = serial(*recalc(STATE["overrides"]))
                return self.sendx(json.dumps(STATE["payload"], ensure_ascii=False).encode(), "application/json; charset=utf-8")
            if path == "/api/importar":
                result = import_overrides(raw)
                return self.sendx(json.dumps(result, ensure_ascii=False).encode(), "application/json; charset=utf-8")
        except Exception as exc:
            return self.sendx(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False).encode(), "application/json; charset=utf-8", 500)
        return self.sendx(b"No encontrado", code=404)

    def log_message(self, *_):
        pass


if __name__ == "__main__":
    print(f"Planeador V2.4: http://{HOST}:{PORT}")
    server = ThreadingHTTPServer((HOST, PORT), H)
    if HOST == "127.0.0.1":
        threading.Timer(0.6, lambda: webbrowser.open(f"http://{HOST}:{PORT}")).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
