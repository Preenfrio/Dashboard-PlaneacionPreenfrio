#!/usr/bin/env python3
"""Lectura segura de cierres semanales para el planeador.

Este módulo no modifica Supabase. Expone:
- obtener_ultimo_cierre(db)
- filtrar_semanas_cerradas(jobs, cierre)
- relojes_iniciales(cierre)

Los timestamptz se convierten a America/Mexico_City y se devuelven como
`datetime` locales sin tzinfo para ser compatibles con el motor actual.
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


def obtener_ultimo_cierre(db):
    """Devuelve el cierre CERRADA más reciente o None."""
    rows = (
        db.table("cierres_semanales")
        .select(
            "anio,semana,estado,ultimo_fin_nelly,disponible_nelly_desde,"
            "ultimo_fin_fortaleza,disponible_fortaleza_desde,"
            "procesos_pendientes,tarimas_pendientes"
        )
        .eq("estado", "CERRADA")
        .order("anio", desc=True)
        .order("semana", desc=True)
        .limit(1)
        .execute()
        .data
    )
    if not rows:
        return None
    row = dict(rows[0])
    row["anio"] = int(row["anio"])
    row["semana"] = int(row["semana"])
    row["disponible_nelly_local"] = _parse_local(row.get("disponible_nelly_desde"))
    row["disponible_fortaleza_local"] = _parse_local(row.get("disponible_fortaleza_desde"))
    return row


def filtrar_semanas_cerradas(jobs, cierre):
    """Quita de la cola activa todos los procesos hasta la semana cerrada.

    La comparación usa (año ISO, semana ISO). El campo `week` se conserva para
    la interfaz, pero la fecha de proceso es la autoridad para el corte.
    """
    if not cierre:
        return list(jobs)
    limite = (int(cierre["anio"]), int(cierre["semana"]))
    activos = []
    for job in jobs:
        process_date = job.get("process_date")
        if not process_date:
            continue
        iso = process_date.isocalendar()
        if (int(iso.year), int(iso.week)) > limite:
            activos.append(job)
    return activos


def relojes_iniciales(cierre):
    """Relojes independientes por cámara para iniciar la cola activa."""
    if not cierre:
        return {"NELLY": None, "FORTALEZA": None}
    return {
        "NELLY": cierre.get("disponible_nelly_local"),
        "FORTALEZA": cierre.get("disponible_fortaleza_local"),
    }


def aplicar_reloj_inicial(jobs, origin, inicio):
    """Adjunta el reloj inicial a trabajos sin alterar su fecha de producción."""
    if inicio is None:
        return jobs
    return [{**job, "clock_floor": inicio} for job in jobs]
