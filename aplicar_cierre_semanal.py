#!/usr/bin/env python3
"""Guarda en Supabase un cierre semanal validado por diagnostico_cierre_semana.py."""
from __future__ import annotations
import argparse
import json
import os
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
from supabase import create_client

ZONA_OPERATIVA = ZoneInfo("America/Mexico_City")

def required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise SystemExit(f"Falta la variable de entorno {name}")
    return value

def local_iso(value: str | None) -> str | None:
    if not value:
        return None
    parsed = datetime.fromisoformat(value)
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=ZONA_OPERATIVA)
    return parsed.isoformat(timespec="minutes")

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("diagnostico")
    parser.add_argument("--aplicar", action="store_true")
    parser.add_argument("--usuario", default="BERNARDO")
    parser.add_argument("--archivo-planeacion", default="PLANEACION_SEMANAL 1.xlsx")
    parser.add_argument("--archivo-movimientos", default="MOVIMIENTOS_SEMANAL.xlsx")
    args = parser.parse_args()

    path = Path(args.diagnostico)
    data = json.loads(path.read_text(encoding="utf-8"))
    semana = int(data["semana"])
    estado = data.get("estado_sugerido")
    counts = data.get("conteos") or {}
    cierres = data.get("cierres_camara") or {}
    nelly = cierres.get("NELLY") or {}
    fortaleza = cierres.get("FORTALEZA") or {}

    errores = []
    if estado != "LISTO_PARA_CIERRE":
        errores.append(f"estado_sugerido={estado}")
    if int(counts.get("REAL_NO_PLANEADO", 0)) != 0:
        errores.append(f"REAL_NO_PLANEADO={counts.get('REAL_NO_PLANEADO')}")
    if not nelly.get("ultimo_fin_real") or not nelly.get("disponible_desde"):
        errores.append("falta cierre real de NELLY")
    if not fortaleza.get("ultimo_fin_real") or not fortaleza.get("disponible_desde"):
        errores.append("falta cierre real de FORTALEZA")
    if errores:
        raise SystemExit("Cierre bloqueado: " + "; ".join(errores))

    ultimo_nelly = local_iso(nelly["ultimo_fin_real"])
    disponible_nelly = local_iso(nelly["disponible_desde"])
    ultimo_fortaleza = local_iso(fortaleza["ultimo_fin_real"])
    disponible_fortaleza = local_iso(fortaleza["disponible_desde"])
    anio = datetime.fromisoformat(ultimo_nelly).year

    payload = {
        "anio": anio,
        "semana": semana,
        "estado": "CERRADA",
        "ultimo_fin_nelly": ultimo_nelly,
        "disponible_nelly_desde": disponible_nelly,
        "ultimo_fin_fortaleza": ultimo_fortaleza,
        "disponible_fortaleza_desde": disponible_fortaleza,
        "procesos_terminados": int(counts.get("TERMINADO", 0)),
        "procesos_terminados_sin_tiempo": int(counts.get("TERMINADO_SIN_TIEMPO_REAL", 0)),
        "procesos_directos": int(counts.get("DIRECTO_SIN_PREENFRIO", 0)),
        "procesos_pendientes": int(counts.get("PLANEACION_SIN_REAL", 0)),
        "tarimas_pendientes": float(data.get("tarimas_pendientes", 0)),
        "archivo_planeacion": args.archivo_planeacion,
        "archivo_movimientos": args.archivo_movimientos,
        "diagnostico_json": data,
        "creado_por": args.usuario,
        "actualizado_en": datetime.now(ZONA_OPERATIVA).isoformat(timespec="seconds"),
    }

    print(json.dumps({"modo": "APLICAR" if args.aplicar else "VISTA_PREVIA", "cierre": payload}, ensure_ascii=False, indent=2))
    if not args.aplicar:
        print("\nVista previa. Para guardar el cierre agrega --aplicar.")
        return

    sb = create_client(required_env("SUPABASE_URL"), required_env("SUPABASE_SERVICE_ROLE_KEY"))
    result = sb.table("cierres_semanales").upsert(payload, on_conflict="anio,semana").execute()
    saved = result.data[0] if result.data else payload
    print("\nCierre semanal guardado correctamente:")
    print(json.dumps(saved, ensure_ascii=False, indent=2, default=str))
    print("\nNo se modificaron procesos ni planificacion manual.")

if __name__ == "__main__":
    main()
