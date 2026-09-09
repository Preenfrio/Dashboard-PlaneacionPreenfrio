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


def norm(v):
    return " ".join(str(v if v is not None else "").strip().upper().split())


def col(df, name, required=True):
    wanted = norm(name)
    for c in df.columns:
        if norm(c) == wanted:
            return c
    if required:
        raise ValueError(f"Falta columna: {name}")
    return None


def as_dt(v):
    d = pd.to_datetime(v, dayfirst=True, errors="coerce")
    return None if pd.isna(d) else d.to_pydatetime()


def hours(v):
    if pd.isna(v) or v == "":
        return 0.0
    if isinstance(v, (datetime, time)):
        return v.hour + v.minute / 60 + v.second / 3600
    if isinstance(v, timedelta):
        return v.total_seconds() / 3600
    if isinstance(v, (int, float)):
        x = float(v)
        return x * 24 if 0 <= x < 1.5 else x
    parts = str(v).strip().split(":")
    try:
        if len(parts) > 1:
            return float(parts[0]) + float(parts[1]) / 60 + (float(parts[2]) / 3600 if len(parts) > 2 else 0)
        return float(parts[0])
    except (TypeError, ValueError):
        return 0.0


def parse_coordinate(value):
    """Acepta decimal o DMS, por ejemplo 14°59'48.9\"N 92°29'52.8\"W."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    text = str(value).strip().upper().replace("–", "-").replace("—", "-")
    if not text:
        return None
    dms = re.findall(r"(\d{1,3})\s*°\s*(\d{1,2})\s*['’]\s*(\d+(?:\.\d+)?)\s*[\"”]?\s*([NSEW])", text)
    if len(dms) >= 2:
        vals = []
        for deg, minute, second, hemi in dms[:2]:
            number = float(deg) + float(minute) / 60 + float(second) / 3600
            if hemi in ("S", "W"):
                number *= -1
            vals.append(number)
        return tuple(vals)
    nums = re.findall(r"[-+]?\d{1,3}(?:\.\d+)", text)
    if len(nums) >= 2:
        lat, lon = float(nums[0]), float(nums[1])
        # Algunas filas omiten el signo W.
        if "W" in text and lon > 0:
            lon *= -1
        if "S" in text and lat > 0:
            lat *= -1
        if -90 <= lat <= 90 and -180 <= lon <= 180:
            return lat, lon
    return None


def haversine_km(a, b):
    if not a or not b:
        return None
    lat1, lon1 = map(math.radians, a)
    lat2, lon2 = map(math.radians, b)
    dlat, dlon = lat2 - lat1, lon2 - lon1
    h = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return 6371.0088 * 2 * math.asin(math.sqrt(h))


def read_locations():
    df = pd.read_excel(EXCEL, sheet_name="UBICACIONES", engine="openpyxl")
    name_col = col(df, "Trazabilidad de finca")
    coord_col = col(df, "Coordenadas")
    locations = {}
    for _, row in df.iterrows():
        name = norm(row[name_col])
        coord = parse_coordinate(row[coord_col])
        if name and coord:
            locations[name] = coord
    precoolers = {
        "NELLY": locations.get("PREENFRIO NELLY"),
        "FORTALEZA": locations.get("PREENFRIO FORTALEZA"),
    }
    return locations, precoolers


def read_catalogue():
    df = pd.read_excel(EXCEL, sheet_name="CATALOGO", engine="openpyxl")
    cols = {n: col(df, n) for n in ["CLIENTE", "DESTINO", "TRÁNSITO (HRS)", "CICLO HORAS FORTALEZA", "CICLO HORAS NELLY", "CRITICIDAD"]}
    result = {}
    for _, row in df.iterrows():
        destination = norm(row[cols["DESTINO"]])
        if not destination:
            continue
        critical_raw = row[cols["CRITICIDAD"]]
        try:
            critical = int(float(critical_raw))
        except (TypeError, ValueError):
            critical = 99
        result[destination] = {
            "client": norm(row[cols["CLIENTE"]]),
            "transit_h": hours(row[cols["TRÁNSITO (HRS)"]]),
            "critical": critical,
            "cycle": {
                "NELLY": hours(row[cols["CICLO HORAS NELLY"]]),
                "FORTALEZA": hours(row[cols["CICLO HORAS FORTALEZA"]]),
            },
        }
    return result


def read_jobs():
    catalogue = read_catalogue()
    locations, precoolers = read_locations()
    df = pd.read_excel(EXCEL, sheet_name="PLANEACION", engine="openpyxl")
    cols = {n: col(df, n) for n in ["Finca", "Destino", "Pallets", "Fecha de proceso", "FECHA CITA"]}
    cols["SEMANA"] = col(df, "SEMANA", False)
    cols["HORAS TRANSITO"] = col(df, "HORAS TRANSITO", False)
    cols["NIVEL CRITICO"] = col(df, "NIVEL CRITICO", False)
    jobs = {}
    for i, row in df.iterrows():
        farm = str(row[cols["Finca"]]).strip()
        destination = str(row[cols["Destino"]]).strip()
        info = catalogue.get(norm(destination))
        process_date = as_dt(row[cols["Fecha de proceso"]])
        appointment = as_dt(row[cols["FECHA CITA"]])
        try:
            pallets = float(row[cols["Pallets"]])
        except (TypeError, ValueError):
            pallets = 0
        if not info or not process_date or not appointment or pallets <= 0:
            continue
        transit_h = hours(row[cols["HORAS TRANSITO"]]) if cols["HORAS TRANSITO"] else info["transit_h"]
        try:
            critical = int(float(row[cols["NIVEL CRITICO"]])) if cols["NIVEL CRITICO"] else info["critical"]
        except (TypeError, ValueError):
            critical = info["critical"]
        try:
            week = str(int(float(row[cols["SEMANA"]]))) if cols["SEMANA"] else str(process_date.isocalendar().week)
        except (TypeError, ValueError):
            week = str(process_date.isocalendar().week)
        farm_coord = locations.get(norm(farm))
        distances = {origin: haversine_km(farm_coord, coord) for origin, coord in precoolers.items()}
        key = f"{norm(farm)}|{norm(destination)}|{process_date.date()}|{appointment.date()}|{info['client']}"
        if key not in jobs:
            jobs[key] = {
                "key": key, "source": i + 2, "farm": farm, "destination": destination,
                "client": info["client"], "process_date": process_date.replace(hour=0, minute=0, second=0, microsecond=0),
                "appointment": appointment, "transit_h": transit_h, "critical": critical,
                "pallets": pallets, "week": week, "cycle_options": info["cycle"],
                "distance_options_km": distances, "location_found": farm_coord is not None,
            }
        else:
            jobs[key]["pallets"] += pallets
    return list(jobs.values())


def allowed(job):
    if job["client"] == "SMART":
        return ["NELLY"] if job["cycle_options"]["NELLY"] > 0 else []
    return [o for o in ("NELLY", "FORTALEZA") if job["cycle_options"][o] > 0]


def pack(origin, sequence):
    if not sequence:
        return [], []
    capacity = CAP[origin]
    current = datetime.combine(sequence[0]["process_date"].date(), time(START_HOUR))
    remaining = {j["key"]: j["pallets"] for j in sequence}
    first, last, cycles = {}, {}, []
    i = 0
    while i < len(sequence):
        space, loads, duration = capacity, [], 0.0
        while space > 1e-9 and i < len(sequence):
            job = sequence[i]
            qty = min(space, remaining[job["key"]])
            loads.append({"key": job["key"], "destination": job["destination"], "pallets": qty})
            duration = max(duration, job["cycle_options"][origin])
            first.setdefault(job["key"], current)
            remaining[job["key"]] -= qty
            space -= qty
            if remaining[job["key"]] <= 1e-9:
                i += 1
        end = current + timedelta(hours=duration)
        cycles.append({"cycle_no": len(cycles) + 1, "origin": origin, "process_date": sequence[0]["process_date"], "start": current, "end": end, "duration_h": duration, "capacity": capacity, "used_pallets": capacity - space, "free_pallets": space, "loads": loads})
        for load in loads:
            if remaining[load["key"]] <= 1e-9:
                last[load["key"]] = end
        current = end + timedelta(minutes=GAP_MIN)
    rows = []
    for order, job in enumerate(sequence, 1):
        end = last[job["key"]]
        arrival = end + timedelta(hours=job["transit_h"])
        slack = (job["appointment"] - arrival).total_seconds() / 3600
        rows.append({**job, "assigned_origin": origin, "assigned_distance_km": job["distance_options_km"].get(origin), "order": order, "start": first[job["key"]], "end": end, "arrival": arrival, "slack_h": slack, "status": "CUMPLE" if slack >= 0 else "NO CUMPLE"})
    return rows, cycles


def automatic(jobs):
    groups = {}
    for job in jobs:
        groups.setdefault(job["process_date"].date(), []).append(job)
    assignments = {}
    for date, group in groups.items():
        queues = {"NELLY": [], "FORTALEZA": []}
        # La prioridad existente se mantiene: cita, criticidad y orden de origen.
        for job in sorted(group, key=lambda x: (x["appointment"], x["critical"], x["source"])):
            best = None
            for origin in allowed(job):
                test_rows, _ = pack(origin, queues[origin] + [job])
                candidate = next(x for x in test_rows if x["key"] == job["key"])
                distance = job["distance_options_km"].get(origin)
                # Jerarquía: 1) cumplir la cita, 2) menor retraso/mayor holgura,
                # 3) menor distancia en línea recta, 4) menos ciclos/carga.
                score = (
                    0 if candidate["status"] == "CUMPLE" else 1,
                    max(0.0, -candidate["slack_h"]),
                    distance if distance is not None else 1_000_000.0,
                    -candidate["slack_h"] if candidate["status"] == "CUMPLE" else 0.0,
                    len(test_rows),
                )
                if best is None or score < best[0]:
                    best = (score, origin)
            if best:
                queues[best[1]].append(job)
        for origin, queue in queues.items():
            for order, job in enumerate(queue, 1):
                assignments[job["key"]] = {"origin": origin, "order": order}
    return assignments


def recalc(overrides=None):
    jobs = read_jobs()
    auto = automatic(jobs)
    ov = overrides if overrides is not None else STATE["overrides"]
    groups = {}
    for job in jobs:
        if job["key"] not in auto:
            continue
        chosen = ov.get(job["key"], auto[job["key"]])
        origin = chosen.get("origin") or auto[job["key"]]["origin"]
        order = int(chosen.get("order") or auto[job["key"]]["order"])
        if origin != "DIRECTO SIN PREENFRIO":
            if job["client"] == "SMART":
                origin = "NELLY"
            if origin not in allowed(job):
                origin = auto[job["key"]]["origin"]
        groups.setdefault((job["process_date"].date(), origin), []).append((order, job))
    rows, cycles = [], []
    for (date, origin), items in sorted(groups.items()):
        sequence = [j for _, j in sorted(items, key=lambda x: (x[0], x[1]["appointment"], x[1]["critical"], x[1]["source"]))]
        if origin == "DIRECTO SIN PREENFRIO":
            for order, job in enumerate(sequence, 1):
                dispatch = datetime.combine(job["process_date"].date(), time(START_HOUR))
                arrival = dispatch + timedelta(hours=job["transit_h"])
                slack = (job["appointment"] - arrival).total_seconds() / 3600
                rows.append({**job, "assigned_origin": origin, "assigned_distance_km": None, "order": order, "start": dispatch, "end": dispatch, "arrival": arrival, "slack_h": slack, "status": "CUMPLE" if slack >= 0 else "NO CUMPLE"})
        else:
            packed_rows, packed_cycles = pack(origin, sequence)
            rows.extend(packed_rows)
            cycles.extend(packed_cycles)
    return rows, cycles


def serial(rows, cycles):
    def iso(v):
        return v.isoformat(timespec="minutes") if isinstance(v, datetime) else v
    row_fields = ["key", "farm", "destination", "client", "process_date", "appointment", "transit_h", "critical", "pallets", "week", "assigned_origin", "assigned_distance_km", "distance_options_km", "location_found", "order", "start", "end", "arrival", "slack_h", "status"]
    cycle_fields = ["cycle_no", "origin", "process_date", "start", "end", "duration_h", "capacity", "used_pallets", "free_pallets", "loads"]
    return {"ok": True, "version": "2.3", "excel": EXCEL.name, "modified": datetime.fromtimestamp(EXCEL.stat().st_mtime).isoformat(timespec="seconds"), "data": [{f: iso(x[f]) for f in row_fields} for x in rows], "cycles": [{f: iso(x[f]) for f in cycle_fields} for x in cycles]}


def payload():
    mtime = EXCEL.stat().st_mtime
    if STATE["payload"] is None or STATE["mtime"] != mtime:
        STATE["mtime"] = mtime
        STATE["overrides"] = {}
        STATE["payload"] = serial(*recalc())
    return STATE["payload"]


def export_bytes():
    p = payload()
    df = pd.DataFrame([{"CLAVE": x["key"], "FECHA_PROCESO": x["process_date"][:10], "CLIENTE": x["client"], "FINCA": x["farm"], "DESTINO": x["destination"], "TARIMAS": x["pallets"], "PREENFRIO_ASIGNADO": x["assigned_origin"], "DISTANCIA_KM": x["assigned_distance_km"], "ORDEN": x["order"], "FECHA_CITA": x["appointment"][:10], "FECHA_INGRESO": x["start"], "FECHA_DESPACHO": x["end"], "ESTADO": x["status"]} for x in p["data"]])
    buffer = io.BytesIO()
    df.to_excel(buffer, sheet_name="PLAN_AJUSTADO", index=False, engine="openpyxl")
    return buffer.getvalue()


def import_overrides(raw):
    df = pd.read_excel(io.BytesIO(raw), sheet_name="PLAN_AJUSTADO", engine="openpyxl")
    cols = {n: col(df, n) for n in ["CLAVE", "PREENFRIO_ASIGNADO", "ORDEN"]}
    ov = {}
    for _, row in df.iterrows():
        ov[str(row[cols["CLAVE"]]).strip()] = {"origin": norm(row[cols["PREENFRIO_ASIGNADO"]]), "order": int(float(row[cols["ORDEN"]]))}
    STATE["overrides"] = ov
    STATE["payload"] = serial(*recalc(ov))
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
            except Exception as e:
                return self.sendx(json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False).encode(), "application/json; charset=utf-8", 500)
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
        except Exception as e:
            return self.sendx(json.dumps({"ok": False, "error": str(e)}, ensure_ascii=False).encode(), "application/json; charset=utf-8", 500)
        return self.sendx(b"No encontrado", code=404)

    def log_message(self, *_):
        pass


if __name__ == "__main__":
    print(f"Planeador V2.3: http://{HOST}:{PORT}")
    server = ThreadingHTTPServer((HOST, PORT), H)
    if HOST == "127.0.0.1":
        threading.Timer(0.6, lambda: webbrowser.open(f"http://{HOST}:{PORT}")).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
