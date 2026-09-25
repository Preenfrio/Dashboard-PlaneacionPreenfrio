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
    if value is None or (isinstance(value, float) and pd.isna(value)) or value == "":
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
        return (
            float(parts[0])
            + (float(parts[1]) / 60 if len(parts) > 1 else 0)
            + (float(parts[2]) / 3600 if len(parts) > 2 else 0)
        )
    except (TypeError, ValueError):
        return 0.0


def parse_coordinate(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    text = str(value).strip().upper().replace("–", "-").replace("—", "-")
    dms = re.findall(
        r'(\d{1,3})\s*°\s*(\d{1,2})\s*[\'’]\s*(\d+(?:\.\d+)?)\s*[\"”]?\s*([NSEW])',
        text,
    )
    if len(dms) >= 2:
        result = []
        for degrees, minutes, seconds, hemisphere in dms[:2]:
            number = float(degrees) + float(minutes) / 60 + float(seconds) / 3600
            result.append(-number if hemisphere in ("S", "W") else number)
        return tuple(result)
    numbers = re.findall(r"[-+]?\d{1,3}(?:\.\d+)", text)
    if len(numbers) >= 2:
        latitude, longitude = float(numbers[0]), float(numbers[1])
        if "W" in text and longitude > 0:
            longitude = -longitude
        if "S" in text and latitude > 0:
            latitude = -latitude
        if -90 <= latitude <= 90 and -180 <= longitude <= 180:
            return latitude, longitude
    return None


def haversine_km(point_a, point_b):
    if not point_a or not point_b:
        return None
    lat1, lon1 = map(math.radians, point_a)
    lat2, lon2 = map(math.radians, point_b)
    dlat, dlon = lat2 - lat1, lon2 - lon1
    haversine = (
        math.sin(dlat / 2) ** 2
        + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    )
    return 6371.0088 * 2 * math.asin(math.sqrt(haversine))


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
    names = [
        "CLIENTE",
        "DESTINO",
        "TRÁNSITO (HRS)",
        "CICLO HORAS FORTALEZA",
        "CICLO HORAS NELLY",
        "CRITICIDAD",
    ]
    columns = {name: col(frame, name) for name in names}
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
    columns = {
        name: col(frame, name)
        for name in ["Finca", "Destino", "Pallets", "Fecha de proceso", "FECHA CITA"]
    }
    for name in ["SEMANA", "HORAS TRANSITO", "NIVEL CRITICO"]:
        columns[name] = col(frame, name, False)

    jobs = []
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

        transit_h = (
            hours(row[columns["HORAS TRANSITO"]])
            if columns["HORAS TRANSITO"]
            else info["transit_h"]
        )
        try:
            critical = (
                int(float(row[columns["NIVEL CRITICO"]]))
                if columns["NIVEL CRITICO"]
                else info["critical"]
            )
        except (TypeError, ValueError):
            critical = info["critical"]
        try:
            week = (
                str(int(float(row[columns["SEMANA"]])))
                if columns["SEMANA"]
                else str(process_date.isocalendar().week)
            )
        except (TypeError, ValueError):
            week = str(process_date.isocalendar().week)

        coordinate = locations.get(norm(farm))
        distances = {
            origin: haversine_km(coordinate, precooler)
            for origin, precooler in precoolers.items()
        }
        key = (
            f"{norm(farm)}|{norm(destination)}|{process_date.date()}|"
            f"{appointment.date()}|{info['client']}|{index + 2}"
        )
        jobs.append(
            {
                "key": key,
                "source": index + 2,
                "farm": farm,
                "destination": destination,
                "client": info["client"],
                "process_date": process_date.replace(hour=0, minute=0, second=0, microsecond=0),
                "appointment": appointment,
                "transit_h": transit_h,
                "critical": critical,
                "pallets": pallets,
                "week": week,
                "cycle_options": info["cycle"],
                "distance_options_km": distances,
                "location_found": coordinate is not None,
            }
        )
    return jobs


def allowed(job):
    farm = norm(job.get("farm"))
    client = norm(job.get("client"))
    options = job.get("cycle_options") or {}
    enabled = {origin: hours(options.get(origin, 0)) > 0 for origin in CAP}

    if farm.startswith("A01 003") or "DOÑA NELLY" in farm or "DONA NELLY" in farm:
        return ["NELLY"] if enabled["NELLY"] else []
    if farm.startswith("A01 002") or farm == "FORTALEZA" or "A01 002 FORTALEZA" in farm:
        return ["FORTALEZA"] if enabled["FORTALEZA"] else []
    if client == "SMART":
        return ["NELLY"] if enabled["NELLY"] else []
    return [origin for origin in ("NELLY", "FORTALEZA") if enabled[origin]]


def release_time(job):
    return datetime.combine(job["process_date"].date(), time(START_HOUR))


def deadline(job):
    return job["appointment"] - timedelta(hours=job["transit_h"])


def daily_priority(job, origin):
    """Orden dentro del mismo día y cámara.

    Si existe orden manual, las flechas dominan la secuencia.
    Sin orden manual, la sugerencia automática prioriza:
    criticidad ascendente, deadline, distancia y fuente.
    """
    distance = job.get("distance_options_km", {}).get(origin)
    distance_value = distance if distance is not None else 1_000_000.0
    manual_order = job.get("manual_order")

    if manual_order is not None:
        return (
            int(manual_order),
            deadline(job),
            job.get("critical", 99),
            distance_value,
            job.get("source", 999999),
        )

    return (
        job.get("critical", 99),
        deadline(job),
        distance_value,
        job.get("source", 999999),
    )


# REGLA_FINCA_LOCAL_V2
# REGLA_RELLENO_FINCA_LOCAL_V3
def is_local_farm_for_origin(job, origin):
    """True cuando el proceso pertenece a la finca donde opera la cámara."""
    farm = norm(job.get("farm"))
    if origin == "NELLY":
        return (
            farm.startswith("A01 003")
            or "DOÑA NELLY" in farm
            or "DONA NELLY" in farm
        )
    if origin == "FORTALEZA":
        return (
            farm.startswith("A01 002")
            or farm == "FORTALEZA"
            or "A01 002 FORTALEZA" in farm
        )
    return False


# REGLA_RELLENO_FINCA_LOCAL_V4
def local_fillers(pending, remaining, current, origin, excluded_keys=None):
    """Fruta de la finca local que puede completar capacidad libre.

    Puede usar procesos locales con fecha de proceso igual o anterior a la
    fecha actual de la cámara. Nunca toma fruta fechada para un día futuro.
    """
    excluded_keys = excluded_keys or set()
    return [
        job
        for job in pending
        if job["key"] not in excluded_keys
        and remaining[job["key"]] > 1e-9
        and job["process_date"].date() <= current.date()
        and is_local_farm_for_origin(job, origin)
    ]


def pack(origin, sequence):
    """FIFO diario, prioridad normal y relleno excepcional con finca local."""
    if not sequence:
        return [], []

    capacity = CAP[origin]
    remaining = {job["key"]: float(job["pallets"]) for job in sequence}
    first, last, portions, cycles = {}, {}, {}, []
    pending = list(sequence)
    current = min(release_time(job) for job in pending)

    while pending:
        available = [
            job
            for job in pending
            if release_time(job) <= current
            and remaining[job["key"]] > 1e-9
        ]

        if not available:
            # Cámara completamente vacía: antes de saltar a las 19:00,
            # puede iniciar fruta de la finca local del mismo día.
            available = local_fillers(
                pending,
                remaining,
                current,
                origin,
            )

        if not available:
            current = min(
                release_time(job)
                for job in pending
                if remaining[job["key"]] > 1e-9
            )
            available = [
                job
                for job in pending
                if release_time(job) <= current
                and remaining[job["key"]] > 1e-9
            ]

        oldest_date = min(job["process_date"].date() for job in available)
        normal_jobs = [
            job
            for job in available
            if job["process_date"].date() == oldest_date
        ]
        normal_jobs.sort(key=lambda job: daily_priority(job, origin))

        # Los procesos normales siempre mandan. Solo si dejan capacidad libre,
        # se agregan procesos de la finca local del día en curso.
        # REGLA_RELLENO_DISPONIBLE_V5
        # Primero manda el día más antiguo. Si deja espacio, se agregan
        # procesos de días posteriores que ya estén liberados a esta hora.
        normal_keys = {job["key"] for job in normal_jobs}
        released_fillers = [
            job
            for job in available
            if job["key"] not in normal_keys
        ]
        released_fillers.sort(
            key=lambda job: (
                job["process_date"].date(),
                daily_priority(job, origin),
            )
        )

        # La excepción de finca local queda como último recurso y únicamente
        # para fruta que todavía no alcanza su liberación normal de las 19:00.
        excluded_keys = normal_keys | {
            job["key"] for job in released_fillers
        }
        early_local_fillers = local_fillers(
            pending,
            remaining,
            current,
            origin,
            excluded_keys=excluded_keys,
        )
        early_local_fillers = [
            job
            for job in early_local_fillers
            if release_time(job) > current
        ]
        early_local_fillers.sort(
            key=lambda job: (
                job["process_date"].date(),
                daily_priority(job, origin),
            )
        )

        cycle_jobs = normal_jobs + released_fillers + early_local_fillers

        space, loads, duration = capacity, [], 0.0
        for job in cycle_jobs:
            if space <= 1e-9:
                break

            quantity = min(space, remaining[job["key"]])
            if quantity <= 1e-9:
                continue

            loads.append(
                {
                    "key": job["key"],
                    "destination": job["destination"],
                    "farm": job.get("farm", ""),
                    "pallets": quantity,
                    "production_date": job["process_date"].date().isoformat(),
                }
            )
            first.setdefault(job["key"], current)
            portions.setdefault(job["key"], []).append(
                {"start": current, "pallets": quantity}
            )
            remaining[job["key"]] -= quantity
            space -= quantity
            duration = max(
                duration,
                hours(job["cycle_options"].get(origin, 0)),
            )

        if not loads:
            raise ValueError(
                f"No se pudo formar ciclo para {origin} el {oldest_date}"
            )

        end = current + timedelta(hours=duration)
        for load in loads:
            portions[load["key"]][-1]["end"] = end

        cycles.append(
            {
                "cycle_no": len(cycles) + 1,
                "origin": origin,
                "process_date": datetime.combine(oldest_date, time()),
                "start": current,
                "end": end,
                "duration_h": duration,
                "capacity": capacity,
                "used_pallets": capacity - space,
                "free_pallets": space,
                "loads": loads,
            }
        )

        completed = []
        for load in loads:
            if remaining[load["key"]] <= 1e-9:
                last[load["key"]] = end
                completed.append(load["key"])

        pending = [
            job for job in pending
            if job["key"] not in completed
        ]
        current = end + timedelta(minutes=GAP_MIN)

    # Mantiene el orden visible real, incluidas las flechas manuales.
    order_by_key = {}
    dates = sorted({job["process_date"].date() for job in sequence})
    for process_date in dates:
        same_day = [
            job
            for job in sequence
            if job["process_date"].date() == process_date
        ]
        same_day.sort(key=lambda job: daily_priority(job, origin))
        for position, job in enumerate(same_day, 1):
            order_by_key[job["key"]] = position

    rows = []
    for job in sequence:
        finish = last[job["key"]]
        dispatch = finish
        arrival = dispatch + timedelta(hours=job["transit_h"])
        slack = (job["appointment"] - arrival).total_seconds() / 3600
        rows.append(
            {
                **job,
                "assigned_origin": origin,
                "assigned_distance_km": job["distance_options_km"].get(origin),
                "order": order_by_key[job["key"]],
                "start": first[job["key"]],
                "end": finish,
                "pre_cooling_end": finish,
                "dispatch": dispatch,
                "arrival": arrival,
                "conservation_wait_h": 0.0,
                "portions": portions[job["key"]],
                "slack_h": slack,
                "status": "CUMPLE" if slack >= 0 else "NO CUMPLE",
            }
        )

    return rows, cycles


def automatic(jobs):
    queues = {"NELLY": [], "FORTALEZA": []}
    assignments = {}
    ordered = sorted(
        jobs,
        key=lambda job: (
            release_time(job),
            job.get("critical", 99),
            deadline(job),
            job["source"],
        ),
    )
    for job in ordered:
        best = None
        for origin in allowed(job):
            rows, cycles = pack(origin, queues[origin] + [job])
            candidate = next(row for row in rows if row["key"] == job["key"])
            score = (
                0 if candidate["status"] == "CUMPLE" else 1,
                max(0, -candidate["slack_h"]),
                len(cycles),
            )
            if best is None or score < best[0]:
                best = (score, origin)
        if best:
            origin = best[1]
            queues[origin].append(job)
            assignments[job["key"]] = {
                "origin": origin,
                "order": 1
                + sum(
                    item["process_date"].date() == job["process_date"].date()
                    for item in queues[origin][:-1]
                ),
            }
    return assignments


def recalc(overrides=None):
    jobs = read_jobs()
    automatic_assignments = automatic(jobs)
    selected_overrides = overrides if overrides is not None else STATE["overrides"]
    queues = {"NELLY": [], "FORTALEZA": []}
    direct = []

    for job in jobs:
        if job["key"] not in automatic_assignments:
            continue
        chosen = selected_overrides.get(job["key"], automatic_assignments[job["key"]])
        origin = chosen.get("origin") or automatic_assignments[job["key"]]["origin"]
        job_copy = {
            **job,
            "manual_order": int(
                chosen.get("order") or automatic_assignments[job["key"]]["order"]
            ),
        }
        if origin == "DIRECTO SIN PREENFRIO":
            direct.append(job_copy)
            continue
        permitted = allowed(job)
        if origin not in permitted:
            automatic_origin = automatic_assignments[job["key"]]["origin"]
            origin = automatic_origin if automatic_origin in permitted else permitted[0]
        queues[origin].append(job_copy)

    rows, cycles = [], []
    for origin in ("NELLY", "FORTALEZA"):
        packed_rows, packed_cycles = pack(origin, queues[origin])
        rows.extend(packed_rows)
        cycles.extend(packed_cycles)

    for order, job in enumerate(
        sorted(direct, key=lambda item: (release_time(item), deadline(item))), 1
    ):
        dispatch = release_time(job)
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
    return rows, cycles


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


def serial(rows, cycles):
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
        "version": "2.5.2",
        "excel": EXCEL.name,
        "modified": datetime.fromtimestamp(EXCEL.stat().st_mtime).isoformat(timespec="seconds"),
        "data": [
            {field: json_value(row.get(field)) for field in row_fields}
            for row in rows
        ],
        "cycles": [
            {field: json_value(cycle.get(field)) for field in cycle_fields}
            for cycle in cycles
        ],
    }


def payload():
    modified_time = EXCEL.stat().st_mtime
    if STATE["payload"] is None or STATE["mtime"] != modified_time:
        STATE.update(mtime=modified_time, overrides={})
        STATE["payload"] = serial(*recalc())
    return STATE["payload"]


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
        try:
            if path == "/":
                return self.sendx(HTML.read_bytes(), "text/html; charset=utf-8")
            if path == "/imagen_portada.jpg":
                image_path = BASE / "imagen_portada.jpg"
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
                    extra={"Content-Disposition": 'attachment; filename="PLAN_AJUSTADO_V2.xlsx"'},
                )
        except Exception as exc:
            body = json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False).encode()
            return self.sendx(body, "application/json; charset=utf-8", 500)
        return self.sendx(b"No encontrado", code=404)

    def do_POST(self):
        try:
            raw = self.rfile.read(int(self.headers.get("Content-Length", "0")))
            if urlparse(self.path).path == "/api/replan":
                STATE["overrides"] = json.loads(raw or b"{}").get("overrides", {})
                STATE["payload"] = serial(*recalc(STATE["overrides"]))
                body = json.dumps(STATE["payload"], ensure_ascii=False, default=str).encode()
                return self.sendx(body, "application/json; charset=utf-8")
        except Exception as exc:
            body = json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False).encode()
            return self.sendx(body, "application/json; charset=utf-8", 500)
        return self.sendx(b"No encontrado", code=404)

    def log_message(self, *_):
        pass


if __name__ == "__main__":
    print(f"Planeador V2.5.2: http://{HOST}:{PORT}")
    server = ThreadingHTTPServer((HOST, PORT), H)
    if HOST == "127.0.0.1":
        threading.Timer(0.6, lambda: webbrowser.open(f"http://{HOST}:{PORT}")).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
