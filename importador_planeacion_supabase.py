#!/usr/bin/env python3
"""Importador inicial corregido de PLANEACION_SEMANAL 1.xlsx a Supabase.

Correcciones:
- No exige columna OV. Si no existe, guarda OV como NULL.
- Distingue filas repetidas mediante ocurrencia calculada en el Excel.
- No sobrescribe tarimas cuando SKU/lote/finca/destino/fecha se repiten.
- Usa Llave_Embarque_Consulta y Ocurrencia_Consulta si existen en el futuro.
- Usa datetime UTC consciente de zona horaria.

Validar sin escribir:
  py importador_planeacion_supabase.py "PLANEACION_SEMANAL 1.xlsx" --dry-run

Importar:
  py importador_planeacion_supabase.py "PLANEACION_SEMANAL 1.xlsx"
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
from datetime import UTC, datetime
from pathlib import Path

import pandas as pd

REQUIRED = [
    "Descripción", "SKU", "Lote", "Finca", "Destino", "Cantidad",
    "Pallets", "Fecha de proceso", "FECHA CITA",
]
OPTIONAL = [
    "SEMANA", "DIAS TRANSITO", "HORAS TRANSITO", "NIVEL CRITICO",
    "OV", "OV OM", "Llave_Base_Consulta", "Ocurrencia_Consulta",
    "Llave_Embarque_Consulta",
]


def norm(value):
    return " ".join(str(value if value is not None else "").strip().upper().split())


def slug(value):
    return re.sub(r"[^A-Z0-9]+", "-", norm(value)).strip("-")


def iso_dt(value):
    parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
    return None if pd.isna(parsed) else parsed.to_pydatetime().isoformat(timespec="seconds")


def number(value):
    try:
        return None if pd.isna(value) else float(value)
    except (TypeError, ValueError):
        return None


def integer(value):
    try:
        return None if pd.isna(value) else int(float(value))
    except (TypeError, ValueError):
        return None


def sha256_file(path):
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sha256_row(data):
    raw = json.dumps(data, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def resolve_columns(frame):
    lookup = {norm(column): column for column in frame.columns}
    missing = [name for name in REQUIRED if norm(name) not in lookup]
    if missing:
        raise ValueError("Faltan columnas: " + ", ".join(missing))
    return {name: lookup.get(norm(name)) for name in REQUIRED + OPTIONAL}


def source_base(record):
    """Grupo base usado únicamente para calcular ocurrencia entre duplicados."""
    return "|".join([
        slug(record["ov"] or "SIN-OV"),
        slug(record["sku"]),
        slug(record["lote"] or "SIN-LOTE"),
        slug(record["finca"]),
        slug(record["destino"]),
        record["fecha_proceso"][:10],
    ])


def build_identity(record):
    """Prioriza llave del origen; si falta, usa grupo base + ocurrencia."""
    source_key = norm(record.get("llave_embarque_consulta"))
    if source_key:
        return "ORIGEN|" + source_key
    return f"{source_base(record)}|OC-{record['ocurrencia_importador']}"


def read_excel(path):
    frame = pd.read_excel(path, sheet_name="PLANEACION", engine="openpyxl")
    columns = resolve_columns(frame)
    records = []
    rejected = []
    occurrence_counter = {}

    # La columna OV no existe hoy en PLANEACION_SEMANAL 1. Se deja NULL.
    ov_column = columns.get("OV") or columns.get("OV OM")

    for index, row in frame.iterrows():
        process_date = iso_dt(row[columns["Fecha de proceso"]])
        appointment = iso_dt(row[columns["FECHA CITA"]])
        pallets = number(row[columns["Pallets"]])

        if pallets is None:
            reason = "Pallets no numérico (por ejemplo, Granel)"
        elif pallets <= 0:
            reason = "Pallets debe ser mayor que cero"
        elif not process_date:
            reason = "Fecha de proceso inválida"
        elif not appointment:
            reason = "Fecha de cita inválida"
        else:
            reason = None

        if reason:
            rejected.append({"fila": index + 2, "motivo": reason})
            continue

        ov = None
        if ov_column and not pd.isna(row[ov_column]):
            candidate = str(row[ov_column]).strip()
            ov = candidate if candidate and norm(candidate) not in {"NAN", "NONE", "NULL"} else None

        record = {
            "ov": ov,
            "descripcion": str(row[columns["Descripción"]]).strip(),
            "sku": str(row[columns["SKU"]]).strip(),
            "lote": None if pd.isna(row[columns["Lote"]]) else str(row[columns["Lote"]]).strip(),
            "finca": str(row[columns["Finca"]]).strip(),
            "destino": str(row[columns["Destino"]]).strip(),
            "cliente": None,
            "cantidad": number(row[columns["Cantidad"]]),
            "pallets": pallets,
            "fecha_proceso": process_date,
            "fecha_cita": appointment,
            "semana": integer(row[columns["SEMANA"]]) if columns.get("SEMANA") else pd.Timestamp(process_date).isocalendar().week,
            "dias_transito": number(row[columns["DIAS TRANSITO"]]) if columns.get("DIAS TRANSITO") else None,
            "horas_transito": number(row[columns["HORAS TRANSITO"]]) if columns.get("HORAS TRANSITO") else None,
            "nivel_critico": integer(row[columns["NIVEL CRITICO"]]) if columns.get("NIVEL CRITICO") else None,
            "llave_base_consulta": None if not columns.get("Llave_Base_Consulta") or pd.isna(row[columns["Llave_Base_Consulta"]]) else str(row[columns["Llave_Base_Consulta"]]).strip(),
            "ocurrencia_consulta": integer(row[columns["Ocurrencia_Consulta"]]) if columns.get("Ocurrencia_Consulta") else None,
            "llave_embarque_consulta": None if not columns.get("Llave_Embarque_Consulta") or pd.isna(row[columns["Llave_Embarque_Consulta"]]) else str(row[columns["Llave_Embarque_Consulta"]]).strip(),
            "fila_excel": index + 2,
        }

        base = source_base(record)
        occurrence_counter[base] = occurrence_counter.get(base, 0) + 1
        # Si el origen ya trae ocurrencia, se respeta; si no, se calcula por orden de filas.
        record["ocurrencia_importador"] = record["ocurrencia_consulta"] or occurrence_counter[base]
        record["identidad_normalizada"] = build_identity(record)
        record["hash_contenido"] = sha256_row({k: v for k, v in record.items() if k != "hash_contenido"})
        records.append(record)

    return records, rejected


def summary_for(path, records, rejected):
    identities = [record["identidad_normalizada"] for record in records]
    return {
        "archivo": path.name,
        "hash": sha256_file(path),
        "filas_validas": len(records),
        "filas_rechazadas": len(rejected),
        "semanas": sorted({record["semana"] for record in records if record["semana"] is not None}),
        "identidades_duplicadas": len(identities) - len(set(identities)),
        "tarimas_validas": round(sum(record["pallets"] for record in records), 3),
        "ov_presentes": sum(bool(record["ov"]) for record in records),
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("excel")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    path = Path(args.excel)
    records, rejected = read_excel(path)
    summary = summary_for(path, records, rejected)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    if rejected:
        print("Rechazos de muestra:")
        print(json.dumps(rejected[:10], ensure_ascii=False, indent=2))
    if summary["identidades_duplicadas"]:
        raise SystemExit("Persisten identidades duplicadas. No se escribirá en Supabase.")
    if args.dry_run:
        return

    from supabase import create_client

    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        raise SystemExit("Faltan SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY")

    supabase = create_client(url, key)
    imported = supabase.table("importaciones").insert({
        "nombre_archivo": path.name,
        "hash_archivo": summary["hash"],
        "semanas": summary["semanas"],
        "filas_leidas": len(records) + len(rejected),
        "filas_validas": len(records),
        "estado": "VALIDADA",
    }).execute().data[0]

    for record in records:
        process = supabase.table("procesos").upsert({
            "identidad_normalizada": record["identidad_normalizada"],
            "activo_fuente": True,
            "estado_sincronizacion": "NUEVO",
            "primera_importacion_id": imported["id"],
            "ultima_importacion_id": imported["id"],
        }, on_conflict="identidad_normalizada").execute().data[0]

        source = {key: value for key, value in record.items() if key not in {
            "identidad_normalizada", "ocurrencia_importador", "fila_excel"
        }}
        source.update({
            "proceso_id": process["id"],
            "archivo_origen": path.name,
            "datos_extra": {
                "fila_excel": record["fila_excel"],
                "ocurrencia_importador": record["ocurrencia_importador"],
            },
        })
        supabase.table("procesos_fuente").upsert(source, on_conflict="proceso_id").execute()
        supabase.table("procesos_planificacion").upsert({
            "proceso_id": process["id"]
        }, on_conflict="proceso_id").execute()

    supabase.table("importaciones").update({
        "estado": "APLICADA",
        "filas_nuevas": len(records),
        "aplicado_en": datetime.now(UTC).isoformat(),
    }).eq("id", imported["id"]).execute()

    print("Importación inicial corregida aplicada.")


if __name__ == "__main__":
    main()
