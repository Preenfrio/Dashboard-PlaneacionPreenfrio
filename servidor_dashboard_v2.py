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


def release_time(job):
    """Hora en que la fruta queda disponible para entrar a preenfrio."""
    return datetime.combine(job["process_date"].date(), time(START_HOUR))


def pack(origin, sequence):
    """Construye una sola cola cronologica por preenfrio, sin reiniciar por dia."""
    if not sequence:
        return [], []
    capacity = CAP[origin]
    sequence = sorted(sequence, key=lambda j: (release_time(j), j.get("manual_order", 9999), j["critical"], j["appointment"], j["source"]))
    remaining = {j["key"]: j["pallets"] for j in sequence}
    first, last, cycles = {}, {}, []
    pending = list(sequence)
    current = min(release_time(j) for j in pending)

    while pending:
        available = [j for j in pending if release_time(j) <= current]
        if not available:
            current = min(release_time(j) for j in pending)
            available = [j for j in pending if release_time(j) <= current]
        available.sort(key=lambda j: (j.get("manual_order", 9999), j["critical"], j["appointment"], j["source"]))

        space, loads, duration = capacity, [], 0.0
        for job in available:
            if space <= 1e-9:
                break
            qty = min(space, remaining[job["key"]])
            if qty <= 1e-9:
                continue
            loads.append({"key": job["key"], "destination": job["destination"], "farm": job.get("farm", ""), "pallets": qty})
            duration = max(duration, job["cycle_options"][origin])
            first.setdefault(job["key"], current)
            remaining[job["key"]] -= qty
            space -= qty

        end = current + timedelta(hours=duration)
        cycles.append({
            "cycle_no": len(cycles) + 1,
            "origin": origin,
            "process_date": current.replace(hour=0, minute=0, second=0, microsecond=0),
            "start": current,
            "end": end,
            "duration_h": duration,
            "capacity": capacity,
            "used_pallets": capacity - space,
            "free_pallets": space,
            "loads": loads,
        })
        completed = []
        for load in loads:
            key = load["key"]
            if remaining[key] <= 1e-9:
                last[key] = end
                completed.append(key)
        pending = [j for j in pending if j["key"] not in completed]
        current = end + timedelta(minutes=GAP_MIN)

    rows = []
    order_by_day = {}
    for job in sequence:
        end = last[job["key"]]
        arrival = end + timedelta(hours=job["transit_h"])
        slack = (job["appointment"] - arrival).total_seconds() / 3600
        day_key = (job["process_date"].date(), origin)
        order_by_day[day_key] = order_by_day.get(day_key, 0) + 1
        rows.append({
            **job,
            "assigned_origin": origin,
            "assigned_distance_km": job["distance_options_km"].get(origin),
            "order": order_by_day[day_key],
            "start": first[job["key"]],
            "end": end,
            "arrival": arrival,
            "slack_h": slack,
            "status": "CUMPLE" if slack >= 0 else "NO CUMPLE",
        })
    return rows, cycles


def _plan_from_queues(queues):
    rows, cycles = [], []
    for origin in ("NELLY", "FORTALEZA"):
        r, c = pack(origin, queues[origin])
        rows.extend(r)
        cycles.extend(c)
    return rows, cycles


def _plan_score(queues):
    """Puntuacion lexicografica: primero se protege la criticidad mas alta.

    Para cada nivel, empezando por 1, se compara:
      1) cantidad de incumplimientos;
      2) horas de retraso acumuladas.
    Solo despues se consideran el total, distancia y numero de ciclos.
    """
    rows, cycles = _plan_from_queues(queues)
    levels = sorted({max(1, int(r.get("critical", 99))) for r in rows})
    # Incluye 1..9 aunque algun nivel no aparezca, para que el criterio sea estable.
    levels = sorted(set(range(1, 10)) | set(levels))
    score = []
    for level in levels:
        same = [r for r in rows if max(1, int(r.get("critical", 99))) == level]
        score.extend([
            sum(r["status"] == "NO CUMPLE" for r in same),
            round(sum(max(0.0, -r["slack_h"]) for r in same), 6),
        ])
    score.extend([
        sum(r["status"] == "NO CUMPLE" for r in rows),
        round(sum(max(0.0, -r["slack_h"]) for r in rows), 6),
        round(sum((r.get("assigned_distance_km") or 0.0) for r in rows), 6),
        len(cycles),
    ])
    return tuple(score)


def _normalise_queues(queues):
    """Mantiene disponibilidad por fecha; dentro del dia prioriza criticidad/cita."""
    out = {"NELLY": list(queues["NELLY"]), "FORTALEZA": list(queues["FORTALEZA"])}
    for origin in out:
        out[origin].sort(key=lambda j: (
            release_time(j),
            j.get("critical", 99),
            j.get("appointment"),
            j.get("source", 999999),
        ))
    return out


def automatic(jobs):
    """Optimiza asignacion y secuencia protegiendo criticidad y citas.

    No pone ciegamente todo nivel 1 antes: evalua el resultado final. Si dar paso
    a otro proceso permite cumplir ambos sin perjudicar niveles superiores, lo usa.
    Si hay conflicto inevitable, evita primero incumplir nivel 1, luego 2, etc.
    """
    ordered = sorted(jobs, key=lambda j: (
        release_time(j), j.get("critical", 99), j["appointment"], j["source"]
    ))
    queues = {"NELLY": [], "FORTALEZA": []}

    # Construccion inicial: para cada proceso prueba todas las camaras permitidas
    # y posiciones factibles dentro del mismo dia.
    for job in ordered:
        best = None
        for origin in allowed(job):
            base = list(queues[origin])
            same_day_positions = [i for i in range(len(base) + 1)
                                  if i == len(base) or base[i]["process_date"].date() >= job["process_date"].date()]
            for pos in same_day_positions:
                trial = {"NELLY": list(queues["NELLY"]), "FORTALEZA": list(queues["FORTALEZA"])}
                trial[origin].insert(pos, job)
                # Nunca adelantar fruta antes de su disponibilidad; pack() tambien lo valida.
                score = _plan_score(trial)
                tie = (score, origin != "NELLY", pos)
                if best is None or tie < best[0]:
                    best = (tie, trial)
        if best is not None:
            queues = best[1]

    # Mejora local: intercambios adyacentes y cambios de camara hasta estabilizar.
    queues = _normalise_queues(queues)
    best_score = _plan_score(queues)
    improved = True
    passes = 0
    while improved and passes < 12:
        improved = False
        passes += 1

        # 1) Intercambios dentro de la misma camara y misma fecha de proceso.
        for origin in ("NELLY", "FORTALEZA"):
            for i in range(len(queues[origin]) - 1):
                a, b = queues[origin][i], queues[origin][i + 1]
                if a["process_date"].date() != b["process_date"].date():
                    continue
                trial = {"NELLY": list(queues["NELLY"]), "FORTALEZA": list(queues["FORTALEZA"])}
                trial[origin][i], trial[origin][i + 1] = trial[origin][i + 1], trial[origin][i]
                score = _plan_score(trial)
                if score < best_score:
                    queues, best_score, improved = trial, score, True
                    break
            if improved:
                break
        if improved:
            continue

        # 2) Mover procesos no SMART a la otra camara y probar posiciones del mismo dia.
        for origin in ("NELLY", "FORTALEZA"):
            other = "FORTALEZA" if origin == "NELLY" else "NELLY"
            for i, job in enumerate(list(queues[origin])):
                if other not in allowed(job):
                    continue
                reduced = {"NELLY": list(queues["NELLY"]), "FORTALEZA": list(queues["FORTALEZA"])}
                reduced[origin].pop(i)
                for pos in range(len(reduced[other]) + 1):
                    trial = {"NELLY": list(reduced["NELLY"]), "FORTALEZA": list(reduced["FORTALEZA"])}
                    trial[other].insert(pos, job)
                    score = _plan_score(trial)
                    if score < best_score:
                        queues, best_score, improved = trial, score, True
                        break
                if improved:
                    break
            if improved:
                break

    # Orden que recibe recalc(). Se numera por dia/camara para conservar las flechas.
    assignments = {}
    for origin in ("NELLY", "FORTALEZA"):
        counters = {}
        for job in queues[origin]:
            day = job["process_date"].date()
            counters[day] = counters.get(day, 0) + 1
            assignments[job["key"]] = {"origin": origin, "order": counters[day]}
    return assignments

def recalc(overrides=None):
    jobs = read_jobs()
    auto = automatic(jobs)
    ov = overrides if overrides is not None else STATE["overrides"]
    queues = {"NELLY": [], "FORTALEZA": []}
    direct = []

    for job in jobs:
        if job["key"] not in auto:
            continue
        chosen = ov.get(job["key"], auto[job["key"]])
        origin = chosen.get("origin") or auto[job["key"]]["origin"]
        manual_order = int(chosen.get("order") or auto[job["key"]]["order"])
        job_copy = {**job, "manual_order": manual_order}
        if origin == "DIRECTO SIN PREENFRIO":
            direct.append(job_copy)
            continue
        if job["client"] == "SMART":
            origin = "NELLY"
        if origin not in allowed(job):
            origin = auto[job["key"]]["origin"]
        queues[origin].append(job_copy)

    rows, cycles = [], []
    for origin in ("NELLY", "FORTALEZA"):
        # Fecha primero; dentro del mismo dia se respeta el orden manual.
        queues[origin].sort(key=lambda j: (release_time(j), j.get("manual_order", 9999), j["appointment"], j["critical"], j["source"]))
        packed_rows, packed_cycles = pack(origin, queues[origin])
        rows.extend(packed_rows)
        cycles.extend(packed_cycles)

    for order, job in enumerate(sorted(direct, key=lambda j: (release_time(j), j.get("manual_order", 9999))), 1):
        dispatch = release_time(job)
        arrival = dispatch + timedelta(hours=job["transit_h"])
        slack = (job["appointment"] - arrival).total_seconds() / 3600
        rows.append({**job, "assigned_origin": "DIRECTO SIN PREENFRIO", "assigned_distance_km": None,
                     "order": order, "start": dispatch, "end": dispatch, "arrival": arrival,
                     "slack_h": slack, "status": "CUMPLE" if slack >= 0 else "NO CUMPLE"})
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
