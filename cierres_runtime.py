#!/usr/bin/env python3
"""Lectura segura de cierres semanales para el planeador.

No modifica Supabase. Separa la cola activa del histórico cerrado y convierte
los timestamptz a la zona operativa America/Mexico_City.
"""
from __future__ import annotations

from datetime import datetime
from zoneinfo import ZoneInfo

ZONA_OPERATIVA = ZoneInfo("America/Mexico_City")


def _parse_local(value):
    if not value:
        return None
    parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=ZONA_OPERATIVA)
    return parsed.astimezone(ZONA_OPERATIVA).replace(tzinfo=None)


def obtener_cierres(db):
    """Devuelve todos los cierres CERRADA ordenados por año y semana."""
    rows = (
        db.table("cierres_semanales")
        .select(
            "id,anio,semana,estado,fecha_cierre,ultimo_fin_nelly,"
            "disponible_nelly_desde,ultimo_fin_fortaleza,"
            "disponible_fortaleza_desde,procesos_terminados,"
            "procesos_terminados_sin_tiempo,procesos_directos,"
            "procesos_pendientes,tarimas_pendientes"
        )
        .eq("estado", "CERRADA")
        .order("anio", desc=False)
        .order("semana", desc=False)
        .execute()
        .data
    )
    result = []
    for source in rows or []:
        row = dict(source)
        row["anio"] = int(row["anio"])
        row["semana"] = int(row["semana"])
        row["ultimo_fin_nelly_local"] = _parse_local(row.get("ultimo_fin_nelly"))
        row["disponible_nelly_local"] = _parse_local(row.get("disponible_nelly_desde"))
        row["ultimo_fin_fortaleza_local"] = _parse_local(row.get("ultimo_fin_fortaleza"))
        row["disponible_fortaleza_local"] = _parse_local(row.get("disponible_fortaleza_desde"))
        result.append(row)
    return result


def obtener_ultimo_cierre(db):
    cierres = obtener_cierres(db)
    return cierres[-1] if cierres else None


def claves_cerradas(cierres):
    return {(int(row["anio"]), int(row["semana"])) for row in cierres or []}


def clave_semana_job(job):
    process_date = job.get("process_date")
    if not process_date:
        return None
    iso = process_date.isocalendar()
    return int(iso.year), int(iso.week)


def es_semana_cerrada(job, cierres):
    return clave_semana_job(job) in claves_cerradas(cierres)


def separar_activos_historicos(jobs, cierres):
    """Separa procesos activos e históricos sin recalcular semanas cerradas."""
    closed = claves_cerradas(cierres)
    active, historical = [], []
    for job in jobs:
        key = clave_semana_job(job)
        if key in closed:
            historical.append(job)
        else:
            active.append(job)
    return active, historical


def filtrar_semanas_cerradas(jobs, cierre):
    """Compatibilidad: conserva solo trabajos posteriores al último cierre."""
    if not cierre:
        return list(jobs)
    limite = (int(cierre["anio"]), int(cierre["semana"]))
    return [job for job in jobs if clave_semana_job(job) and clave_semana_job(job) > limite]


def relojes_iniciales(cierre):
    if not cierre:
        return {"NELLY": None, "FORTALEZA": None}
    return {
        "NELLY": cierre.get("disponible_nelly_local"),
        "FORTALEZA": cierre.get("disponible_fortaleza_local"),
    }


def cierre_publico(row):
    """Versión serializable y segura para el navegador."""
    if not row:
        return None
    fields = (
        "id", "anio", "semana", "estado", "fecha_cierre",
        "ultimo_fin_nelly_local", "disponible_nelly_local",
        "ultimo_fin_fortaleza_local", "disponible_fortaleza_local",
        "procesos_terminados", "procesos_terminados_sin_tiempo",
        "procesos_directos", "procesos_pendientes", "tarimas_pendientes",
    )
    return {field: row.get(field) for field in fields}
