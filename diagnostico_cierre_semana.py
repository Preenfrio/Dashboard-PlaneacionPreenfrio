#!/usr/bin/env python3
"""Diagnóstico seguro de cierre semanal: Planeación vs operación real.

- No modifica Excel.
- No escribe en Supabase.
- Genera CSV/JSON con cruce, pendientes, parciales y reloj real por cámara.
"""
from __future__ import annotations

import argparse
import json
import math
import re
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import pandas as pd


def text(v: Any) -> str:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    s = unicodedata.normalize("NFKD", str(v).strip().upper())
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(re.sub(r"[^A-Z0-9]+", " ", s).split())


def slug(v: Any) -> str:
    return text(v).replace(" ", "-")


def number(v: Any) -> float | None:
    try:
        if v is None or pd.isna(v):
            return None
        s = str(v).strip().replace(",", "")
        if not s or text(s) in {"N A", "NA", "PDT", "SIN PRE", "SIN PREENFRIO", "GRANEL"}:
            return None
        return float(s)
    except Exception:
        return None


def date_value(v: Any) -> datetime | None:
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)) and 1 <= float(v) <= 100000:
        return datetime(1899, 12, 30) + timedelta(days=float(v))
    for dayfirst in (False, True):
        x = pd.to_datetime(v, errors="coerce", dayfirst=dayfirst)
        if not pd.isna(x):
            return x.to_pydatetime()
    return None


def iso_date(v: Any) -> str | None:
    x = date_value(v)
    return x.date().isoformat() if x else None


def time_delta(v: Any) -> timedelta | None:
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, timedelta):
        return v
    if isinstance(v, (int, float)):
        return timedelta(days=float(v) % 1)
    s = str(v).strip()
    m = re.match(r"^(\d{1,3}):(\d{1,2})(?::(\d{1,2}))?$", s)
    if m:
        h, mi, sec = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
        return timedelta(hours=h, minutes=mi, seconds=sec)
    x = pd.to_datetime(s, errors="coerce")
    return None if pd.isna(x) else timedelta(hours=x.hour, minutes=x.minute, seconds=x.second)


def combine(d: Any, t: Any) -> datetime | None:
    base = date_value(d)
    if not base:
        return None
    delta = time_delta(t) or timedelta(0)
    return base.replace(hour=0, minute=0, second=0, microsecond=0) + delta


def col(df: pd.DataFrame, *names: str) -> str | None:
    lookup = {text(c): c for c in df.columns}
    for n in names:
        if text(n) in lookup:
            return lookup[text(n)]
    return None


def farm(v: Any) -> str:
    s = text(v)
    aliases = {
        "FN CHI DONA NELLY": "A01 003 DONA NELLY",
        "FN CHI MAZATAN": "A11 045 SECTOR MAZATAN",
        "FN CHI FORTALEZA": "A01 002 FORTALEZA",
    }
    return aliases.get(s, s)


def base_key(r: dict[str, Any]) -> str:
    return "|".join([
        slug(r.get("sku")), slug(r.get("lote")), slug(r.get("finca")),
        slug(r.get("destino")), (r.get("fecha_proceso") or "").replace("-", "")
    ])


def occurrence_sort(r: dict[str, Any]) -> tuple:
    return (
        r.get("fecha_cita") or "9999-12-31",
        r.get("cantidad") if r.get("cantidad") is not None else 10**15,
        r.get("pallets") if r.get("pallets") is not None else 10**15,
        r.get("fila_excel", 10**9),
    )


def assign_occurrences(rows: list[dict[str, Any]]) -> None:
    groups: dict[str, list[dict[str, Any]]] = {}
    for r in rows:
        groups.setdefault(r["llave_base"], []).append(r)
    for group in groups.values():
        group.sort(key=occurrence_sort)
        for i, r in enumerate(group, 1):
            r["ocurrencia"] = i
            r["llave_cruce"] = f"{r['llave_base']}|{i}"


def normalise_preenfrio(v: Any) -> str:
    s = text(v)
    if "DIRECTO" in s or s in {"SIN PRE", "SIN PREENFRIO"}:
        return "DIRECTO SIN PREENFRIO"
    if "NELLY" in s:
        return "NELLY"
    if "FORTALEZA" in s:
        return "FORTALEZA"
    return s or "SIN DATO"


def read_planning(path: Path, week: int) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    df = pd.read_excel(path, sheet_name="PLANEACION", engine="openpyxl")
    c = {k: col(df, *v) for k, v in {
        "sku": ("SKU",), "lote": ("Lote",), "finca": ("Finca",), "destino": ("Destino",),
        "cantidad": ("Cantidad",), "pallets": ("Pallets",), "fecha_proceso": ("Fecha de proceso",),
        "semana": ("SEMANA",), "fecha_cita": ("FECHA CITA",), "descripcion": ("Descripción", "Descripcion")
    }.items()}
    missing = [k for k in ("sku", "lote", "finca", "destino", "fecha_proceso", "semana") if not c[k]]
    if missing: raise ValueError("PLANEACION: faltan columnas: " + ", ".join(missing))
    out, rejected = [], []
    for idx, row in df.iterrows():
        w = number(row.get(c["semana"]))
        if w is None or int(w) != week: continue
        r = {
            "origen": "PLANEACION", "fila_excel": idx + 2,
            "descripcion": row.get(c["descripcion"]) if c["descripcion"] else None,
            "sku": text(row.get(c["sku"])), "lote": text(row.get(c["lote"])),
            "finca": farm(row.get(c["finca"])), "destino": text(row.get(c["destino"])),
            "cantidad": number(row.get(c["cantidad"])) if c["cantidad"] else None,
            "pallets": number(row.get(c["pallets"])) if c["pallets"] else None,
            "fecha_proceso": iso_date(row.get(c["fecha_proceso"])),
            "fecha_cita": iso_date(row.get(c["fecha_cita"])) if c["fecha_cita"] else None,
            "semana": week,
        }
        reasons = [f"{k} vacío" for k in ("sku", "lote", "finca", "destino", "fecha_proceso") if not r[k]]
        if r["pallets"] is None: reasons.append("pallets no numérico")
        if reasons: rejected.append({**r, "motivo_rechazo": "; ".join(reasons)}); continue
        r["llave_base"] = base_key(r); out.append(r)
    assign_occurrences(out)
    return out, rejected


def read_real(path: Path, week: int) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    df = pd.read_excel(path, sheet_name="Hoja1", engine="openpyxl")
    names = {
        "sku": ("SKU",), "lote": ("Lote",), "finca": ("Nombre de Organización", "Nombre de Organizacion", "Finca"),
        "destino": ("Destino",), "cantidad": ("Cantidad",), "pallets": ("Pallets",),
        "fecha_proceso": ("Fecha de proceso",), "semana": ("SEMANA",), "fecha_cita": ("FECHA CITA",),
        "preenfrio": ("PREENFRIO",), "fecha_ingreso": ("FECHA INGRESO AL PREENFRIO",),
        "hora_ingreso": ("HORA INGRESO AL PREENFRIO",), "horas_ciclo": ("HORAS DE CICLO",),
        "fecha_despacho": ("FECHA DE DESPACHO REAL",), "hora_despacho": ("HORA DE DESPACHO REAL",),
        "cumplimiento": ("CUMPLIMIENTO",), "motivo": ("MOTIVO",), "observaciones": ("OBSERVACIONES",),
        "folio": ("FOLIO DE DESPACHO",), "descripcion": ("Descripción", "Descripcion"),
    }
    c = {k: col(df, *v) for k, v in names.items()}
    missing = [k for k in ("sku", "lote", "finca", "destino", "fecha_proceso", "semana") if not c[k]]
    if missing: raise ValueError("MOVIMIENTOS: faltan columnas: " + ", ".join(missing))
    out, rejected = [], []
    for idx, row in df.iterrows():
        w = number(row.get(c["semana"]));
        if w is None or int(w) != week: continue
        inicio = combine(row.get(c["fecha_ingreso"]), row.get(c["hora_ingreso"])) if c["fecha_ingreso"] else None
        dur = time_delta(row.get(c["horas_ciclo"])) if c["horas_ciclo"] else None
        fin = inicio + dur if inicio and dur else None
        despacho = combine(row.get(c["fecha_despacho"]), row.get(c["hora_despacho"])) if c["fecha_despacho"] else None
        r = {
            "origen": "MOVIMIENTOS", "fila_excel": idx + 2,
            "descripcion": row.get(c["descripcion"]) if c["descripcion"] else None,
            "sku": text(row.get(c["sku"])), "lote": text(row.get(c["lote"])),
            "finca": farm(row.get(c["finca"])), "destino": text(row.get(c["destino"])),
            "cantidad": number(row.get(c["cantidad"])) if c["cantidad"] else None,
            "pallets": number(row.get(c["pallets"])) if c["pallets"] else None,
            "fecha_proceso": iso_date(row.get(c["fecha_proceso"])),
            "fecha_cita": iso_date(row.get(c["fecha_cita"])) if c["fecha_cita"] else None,
            "semana": week, "preenfrio_real": normalise_preenfrio(row.get(c["preenfrio"])) if c["preenfrio"] else "SIN DATO",
            "inicio_real": inicio.isoformat(timespec="minutes") if inicio else None,
            "fin_preenfrio_real": fin.isoformat(timespec="minutes") if fin else None,
            "despacho_real": despacho.isoformat(timespec="minutes") if despacho else None,
            "horas_ciclo": round(dur.total_seconds()/3600, 4) if dur else None,
            "cumplimiento_real": text(row.get(c["cumplimiento"])) if c["cumplimiento"] else None,
            "motivo": text(row.get(c["motivo"])) if c["motivo"] else None,
            "observaciones": str(row.get(c["observaciones"])).strip() if c["observaciones"] and not pd.isna(row.get(c["observaciones"])) else None,
            "folio_despacho": str(row.get(c["folio"])).strip() if c["folio"] and not pd.isna(row.get(c["folio"])) else None,
        }
        reasons = [f"{k} vacío" for k in ("sku", "lote", "finca", "destino", "fecha_proceso") if not r[k]]
        if r["pallets"] is None: reasons.append("pallets no numérico")
        if reasons: rejected.append({**r, "motivo_rechazo": "; ".join(reasons)}); continue
        r["llave_base"] = base_key(r); out.append(r)
    assign_occurrences(out)
    return out, rejected


def close_diagnostic(planning: list[dict[str, Any]], real: list[dict[str, Any]], week: int) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    pmap = {r["llave_cruce"]: r for r in planning}; rmap = {r["llave_cruce"]: r for r in real}
    detail = []
    for key in sorted(set(pmap) | set(rmap)):
        p, r = pmap.get(key), rmap.get(key)
        if p and r:
            pp, rp = p.get("pallets") or 0, r.get("pallets") or 0
            if rp >= pp - 1e-6:
                # Si existe movimiento pero el área operativa ya no puede
                # recuperar ingreso/ciclo, el proceso conserva la planeación.
                # No se inventa un tiempo real ni se bloquea el cierre.
                status = "TERMINADO" if r.get("fin_preenfrio_real") or r.get("preenfrio_real") == "DIRECTO SIN PREENFRIO" else "TERMINADO_SIN_TIEMPO_REAL"
            elif rp > 0:
                status = "PARCIAL"
            else:
                status = "PENDIENTE"
        elif p:
            # Sin movimiento real: conservar íntegramente la planeación.
            status = "PLANEACION_SIN_REAL"
        else:
            status = "REAL_NO_PLANEADO"
        detail.append({
            "estado_cierre": status, "llave_cruce": key,
            "fila_planeacion": p.get("fila_excel") if p else None,
            "fila_movimientos": r.get("fila_excel") if r else None,
            "semana": week, "fecha_proceso": (r or p).get("fecha_proceso"),
            "sku": (r or p).get("sku"), "lote": (r or p).get("lote"),
            "finca": (r or p).get("finca"), "destino": (r or p).get("destino"),
            "pallets_planeadas": p.get("pallets") if p else None,
            "pallets_reales": r.get("pallets") if r else None,
            "pallets_pendientes": max(
                0,
                (p.get("pallets") or 0) - ((r or {}).get("pallets") or 0),
            ) if p else None,
            "cantidad_planeada": p.get("cantidad") if p else None,
            "cantidad_real": r.get("cantidad") if r else None,
            "preenfrio_real": r.get("preenfrio_real") if r else None,
            "inicio_real": r.get("inicio_real") if r else None,
            "fin_preenfrio_real": r.get("fin_preenfrio_real") if r else None,
            "despacho_real": r.get("despacho_real") if r else None,
            "cumplimiento_real": r.get("cumplimiento_real") if r else None,
            "motivo": r.get("motivo") if r else None,
            "observaciones": r.get("observaciones") if r else None,
            "folio_despacho": r.get("folio_despacho") if r else None,
        })

    camera_closes = {}
    for camera in ("NELLY", "FORTALEZA"):
        finishes = [datetime.fromisoformat(r["fin_preenfrio_real"]) for r in real if r.get("preenfrio_real") == camera and r.get("fin_preenfrio_real")]
        last = max(finishes) if finishes else None
        camera_closes[camera] = {
            "ultimo_fin_real": last.isoformat(timespec="minutes") if last else None,
            "disponible_desde": (last + timedelta(minutes=30)).isoformat(timespec="minutes") if last else None,
            "movimientos_con_fin_real": len(finishes),
        }

    counts: dict[str, int] = {}
    for r in detail: counts[r["estado_cierre"]] = counts.get(r["estado_cierre"], 0) + 1
    summary = {
        "semana": week,
        # Solo un movimiento real no relacionado bloquea el cierre. Los
        # faltantes reales se preservan mediante la planeación.
        "estado_sugerido": "LISTO_PARA_CIERRE" if counts.get("REAL_NO_PLANEADO", 0) == 0 else "REQUIERE_REVISION",
        "conteos": counts,
        "tarimas_planeadas": round(sum(r.get("pallets") or 0 for r in planning), 6),
        "tarimas_reales": round(sum(r.get("pallets") or 0 for r in real), 6),
        "tarimas_pendientes": round(sum(r.get("pallets_pendientes") or 0 for r in detail), 6),
        "cierres_camara": camera_closes,
    }
    return summary, detail


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("planeacion")
    parser.add_argument("movimientos")
    parser.add_argument("--semana", type=int, default=39)
    parser.add_argument("--salida", default=None)
    args = parser.parse_args()
    base_name = args.salida or f"diagnostico_cierre_semana_{args.semana}"
    planning, prejected = read_planning(Path(args.planeacion), args.semana)
    real, rrejected = read_real(Path(args.movimientos), args.semana)
    summary, detail = close_diagnostic(planning, real, args.semana)

    csv_path = Path(base_name + ".csv")
    json_path = Path(base_name + ".json")
    rejected_path = Path(base_name + "_rechazados.csv")
    pd.DataFrame(detail).to_csv(csv_path, index=False, encoding="utf-8-sig")
    pd.DataFrame(prejected + rrejected).to_csv(rejected_path, index=False, encoding="utf-8-sig")
    payload = {
        "archivo_planeacion": Path(args.planeacion).name,
        "archivo_movimientos": Path(args.movimientos).name,
        "filas_validas_planeacion": len(planning),
        "filas_validas_movimientos": len(real),
        "filas_rechazadas_planeacion": len(prejected),
        "filas_rechazadas_movimientos": len(rrejected),
        **summary,
        "archivos_generados": [str(csv_path), str(rejected_path)],
        "modo": "SOLO_DIAGNOSTICO_SIN_ESCRITURA",
    }
    json_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    print("\nNo se modifico ningun Excel ni se escribio nada en Supabase.")


if __name__ == "__main__":
    main()
