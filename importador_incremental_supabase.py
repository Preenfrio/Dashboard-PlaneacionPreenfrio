#!/usr/bin/env python3
"""Importación incremental de PLANEACION_SEMANAL 1.xlsx a Supabase.

Vista previa, no escribe:
  py importador_incremental_supabase.py "PLANEACION_SEMANAL 1.xlsx"

Aplicar cambios:
  py importador_incremental_supabase.py "PLANEACION_SEMANAL 1.xlsx" --aplicar

Reglas:
- Actualiza datos de origen sin borrar planeación manual.
- Reconoce primero llaves exactas; después OV+SKU y SKU+lote.
- Marca retirados solo dentro de las semanas incluidas en el Excel importado.
- Los casos dudosos se registran como ambiguos y no se aplican automáticamente.
"""
from __future__ import annotations
import argparse, hashlib, json, os, re
from datetime import UTC, datetime
from pathlib import Path
import pandas as pd

REQUIRED=["Descripción","SKU","Lote","Finca","Destino","Cantidad","Pallets","Fecha de proceso","FECHA CITA"]
OPTIONAL=["SEMANA","DIAS TRANSITO","HORAS TRANSITO","NIVEL CRITICO","OV","OV OM","Llave_Base_Consulta","Ocurrencia_Consulta","Llave_Embarque_Consulta"]
SOURCE_FIELDS=["llave_base_consulta","ocurrencia_consulta","llave_embarque_consulta","ov","descripcion","sku","lote","finca","destino","cliente","cantidad","pallets","fecha_proceso","semana","fecha_cita","dias_transito","horas_transito","nivel_critico","hash_contenido"]

def norm(v): return " ".join(str(v if v is not None else "").strip().upper().split())
def slug(v): return re.sub(r"[^A-Z0-9]+","-",norm(v)).strip("-")
def dt(v):
    x=pd.to_datetime(v,dayfirst=True,errors="coerce")
    return None if pd.isna(x) else x.to_pydatetime().isoformat(timespec="seconds")
def num(v):
    try:return None if pd.isna(v) else float(v)
    except:return None
def integer(v):
    try:return None if pd.isna(v) else int(float(v))
    except:return None
def hash_json(v):return hashlib.sha256(json.dumps(v,ensure_ascii=False,sort_keys=True,default=str).encode()).hexdigest()
def hash_file(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def cols(df):
    m={norm(c):c for c in df.columns}; missing=[x for x in REQUIRED if norm(x) not in m]
    if missing:raise ValueError("Faltan columnas: "+", ".join(missing))
    return {x:m.get(norm(x)) for x in REQUIRED+OPTIONAL}
def valid_ov(v):
    """Acepta OV reales; descarta 0, vacios y mensajes de diagnostico."""
    s = norm(v)
    if not s or s in {"0", "NAN", "NONE"}:
        return None
    bad = (
        "OV ND", "SIN OV", "NO COMPATIBLE", "NO ENCONTRADA",
        "NO ENCONTRADO", "PENDIENTE", "ERROR"
    )
    if any(x in s for x in bad):
        return None
    return s

def canonical_key(r):
    """Identidad estable basada en datos de negocio, no en texto formateado de Power Query."""
    fp = (r.get("fecha_proceso") or "")[:10].replace("-", "")
    fc = (r.get("fecha_cita") or "")[:10].replace("-", "")
    occurrence = r.get("ocurrencia_consulta") or r.get("ocurrencia_importador") or 1
    return "CANON|" + "|".join([
        slug(r.get("finca") or ""),
        fp,
        fc,
        slug(r.get("destino") or ""),
        slug(r.get("sku") or ""),
        slug(r.get("lote") or ""),
        str(int(occurrence)),
    ])

def base_key(r):
    return "|".join([
        slug(valid_ov(r.get("ov")) or "SIN-OV"),
        slug(r["sku"]),
        slug(r.get("lote") or "SIN-LOTE"),
        slug(r["finca"]),
        slug(r["destino"]),
        r["fecha_proceso"][:10],
    ])

def identity(r):
    # Desde esta version la identidad nueva usa la llave canonica.
    # La conciliacion tambien crea alias canonicos para registros historicos.
    return canonical_key(r)

def read_excel(path):
    df=pd.read_excel(path,sheet_name="PLANEACION",engine="openpyxl"); c=cols(df); out=[]; rejected=[]; seen={}; ovcol=c.get("OV") or c.get("OV OM")
    for i,row in df.iterrows():
        fp=dt(row[c["Fecha de proceso"]]); fc=dt(row[c["FECHA CITA"]]); pal=num(row[c["Pallets"]])
        reason="Pallets no numérico (por ejemplo, Granel)" if pal is None else "Pallets debe ser mayor que cero" if pal<=0 else "Fecha de proceso inválida" if not fp else "Fecha de cita inválida" if not fc else None
        if reason: rejected.append({"fila":i+2,"motivo":reason}); continue
        ov=None if not ovcol or pd.isna(row[ovcol]) else valid_ov(row[ovcol])
        r={"ov":ov,"descripcion":str(row[c["Descripción"]]).strip(),"sku":str(row[c["SKU"]]).strip(),"lote":None if pd.isna(row[c["Lote"]]) else str(row[c["Lote"]]).strip(),"finca":str(row[c["Finca"]]).strip(),"destino":str(row[c["Destino"]]).strip(),"cliente":None,"cantidad":num(row[c["Cantidad"]]),"pallets":pal,"fecha_proceso":fp,"fecha_cita":fc,"semana":integer(row[c["SEMANA"]]) if c.get("SEMANA") else pd.Timestamp(fp).isocalendar().week,"dias_transito":num(row[c["DIAS TRANSITO"]]) if c.get("DIAS TRANSITO") else None,"horas_transito":num(row[c["HORAS TRANSITO"]]) if c.get("HORAS TRANSITO") else None,"nivel_critico":integer(row[c["NIVEL CRITICO"]]) if c.get("NIVEL CRITICO") else None,"llave_base_consulta":None if not c.get("Llave_Base_Consulta") or pd.isna(row[c["Llave_Base_Consulta"]]) else str(row[c["Llave_Base_Consulta"]]).strip(),"ocurrencia_consulta":integer(row[c["Ocurrencia_Consulta"]]) if c.get("Ocurrencia_Consulta") else None,"llave_embarque_consulta":None if not c.get("Llave_Embarque_Consulta") or pd.isna(row[c["Llave_Embarque_Consulta"]]) else str(row[c["Llave_Embarque_Consulta"]]).strip(),"fila_excel":i+2}
        b=base_key(r);seen[b]=seen.get(b,0)+1;r["ocurrencia_importador"]=r["ocurrencia_consulta"] or seen[b];r["identidad_normalizada"]=identity(r)
        r["hash_contenido"]=hash_json({k:v for k,v in r.items() if k not in ("fila_excel","hash_contenido","identidad_normalizada")});out.append(r)
    return out,rejected

def daydiff(a,b):
    try:return abs((pd.Timestamp(a)-pd.Timestamp(b)).total_seconds())/86400
    except:return 9999

def similarity(new, old):
    """Puntaje de conciliacion para distinguir partidas repetidas sin adivinar."""
    score = 0

    new_ov = valid_ov(new.get("ov"))
    old_ov = valid_ov(old.get("ov"))
    if new_ov and new_ov == old_ov:
        score += 70

    if norm(new.get("sku")) == norm(old.get("sku")):
        score += 30
    if norm(new.get("lote")) and norm(new.get("lote")) == norm(old.get("lote")):
        score += 35
    if norm(new.get("finca")) == norm(old.get("finca")):
        score += 15
    if norm(new.get("destino")) == norm(old.get("destino")):
        score += 20

    # Las dos fechas exactas distinguen embarques del mismo lote/SKU.
    if daydiff(new.get("fecha_proceso"), old.get("fecha_proceso")) == 0:
        score += 25
    elif daydiff(new.get("fecha_proceso"), old.get("fecha_proceso")) <= 1:
        score += 8

    if daydiff(new.get("fecha_cita"), old.get("fecha_cita")) == 0:
        score += 35
    elif daydiff(new.get("fecha_cita"), old.get("fecha_cita")) <= 1:
        score += 10

    if abs((new.get("pallets") or 0) - (old.get("pallets") or 0)) < 0.000001:
        score += 20
    if abs((new.get("cantidad") or 0) - (old.get("cantidad") or 0)) < 0.000001:
        score += 15

    new_occ = new.get("ocurrencia_consulta") or new.get("ocurrencia_importador")
    old_occ = old.get("ocurrencia_consulta")
    if new_occ is not None and old_occ is not None and int(new_occ) == int(old_occ):
        score += 25

    # Si la llave base coincide exactamente, es una señal fuerte aunque la llave
    # de embarque haya cambiado de formato por nombres de dias.
    if norm(new.get("llave_base_consulta")) and norm(new.get("llave_base_consulta")) == norm(old.get("llave_base_consulta")):
        score += 45

    return score

def comparable_value(field, value):
    """Normaliza únicamente los campos de negocio guardados en procesos_fuente."""
    if field in {"cantidad", "pallets", "dias_transito", "horas_transito"}:
        return None if value is None else round(float(value), 6)
    if field in {"semana", "nivel_critico", "ocurrencia_consulta"}:
        return None if value is None else int(value)
    if field in {"fecha_proceso", "fecha_cita"}:
        parsed = pd.to_datetime(value, errors="coerce")
        return None if pd.isna(parsed) else parsed.to_pydatetime().isoformat(timespec="seconds")
    return None if value is None else norm(value)


def business_changes(new, old):
    """Devuelve solo diferencias reales; ignora hash, fila Excel y metadatos."""
    fields = [
        "llave_base_consulta", "ocurrencia_consulta", "llave_embarque_consulta",
        "ov", "descripcion", "sku", "lote", "finca", "destino", "cliente",
        "cantidad", "pallets", "fecha_proceso", "semana", "fecha_cita",
        "dias_transito", "horas_transito", "nivel_critico",
    ]
    return {
        field: {"antes": old.get(field), "despues": new.get(field)}
        for field in fields
        if comparable_value(field, old.get(field)) != comparable_value(field, new.get(field))
    }


def strict_base_key(r):
    """Mismo proceso productivo sin usar OV ni fecha de cita, que sí pueden cambiar."""
    fp = (r.get("fecha_proceso") or "")[:10]
    return "|".join([
        norm(r.get("sku")),
        norm(r.get("lote")),
        norm(r.get("finca")),
        norm(r.get("destino")),
        fp,
    ])

def classify(records, existing):
    """
    Conciliación conservadora:
    1) llave canónica / llave de embarque / identidad exacta;
    2) misma base productiva: SKU+lote+finca+destino+fecha de proceso;
    3) desempate por OV, cantidad+pallets, cita u ocurrencia;
    4) jamás cruza fincas, destinos o lotes solo porque coincida la OV.
    """
    def add_index(index, key, row):
        if key:
            index.setdefault(key, []).append(row)

    by_identity = {}
    by_source = {}
    by_canonical = {}
    by_base = {}
    for old in existing:
        add_index(by_identity, old.get("identidad_normalizada"), old)
        add_index(by_source, norm(old.get("llave_embarque_consulta")), old)
        add_index(by_canonical, canonical_key(old), old)
        add_index(by_base, strict_base_key(old), old)

    used = set()
    result = {"nuevos": [], "modificados": [], "sin_cambios": [], "ambiguos": [], "retirados": []}

    def unused(rows):
        return [x for x in rows if x["proceso_id"] not in used]

    def unique(rows):
        rows = unused(rows)
        return rows[0] if len(rows) == 1 else None

    for r in records:
        match = None
        method = None

        exact_tiers = [
            (by_canonical.get(canonical_key(r), []), "CANONICA"),
            (by_source.get(norm(r.get("llave_embarque_consulta")), []), "LLAVE_EMBARQUE"),
            (by_identity.get(r.get("identidad_normalizada"), []), "IDENTIDAD"),
        ]
        for rows, name in exact_tiers:
            match = unique(rows)
            if match:
                method = name
                break

        if not match:
            candidates = unused(by_base.get(strict_base_key(r), []))

            if candidates:
                # 1. OV exacta real.
                ov = valid_ov(r.get("ov"))
                if ov:
                    rows = [x for x in candidates if valid_ov(x.get("ov")) == ov]
                    if len(rows) == 1:
                        match, method = rows[0], "BASE+OV"

                # 2. Cantidad y pallets exactos.
                if not match:
                    rows = [x for x in candidates
                            if abs((r.get("cantidad") or 0) - (x.get("cantidad") or 0)) < 0.000001
                            and abs((r.get("pallets") or 0) - (x.get("pallets") or 0)) < 0.000001]
                    if len(rows) == 1:
                        match, method = rows[0], "BASE+CANTIDAD"

                # 3. Fecha de cita exacta.
                if not match:
                    rows = [x for x in candidates if daydiff(r.get("fecha_cita"), x.get("fecha_cita")) == 0]
                    if len(rows) == 1:
                        match, method = rows[0], "BASE+CITA"

                # 4. Si solo queda un registro con la misma base, es el mismo proceso modificado.
                if not match and len(candidates) == 1:
                    match, method = candidates[0], "BASE_UNICA"

                if not match:
                    # Si existen varias partidas historicas con la misma base,
                    # pero ninguna comparte OV real, cantidad+pallets o cita exacta,
                    # no debemos adivinar: es una partida adicional NUEVA.
                    # Los historicos que desaparecieron se evaluaran como retirados.
                    candidate_ovs = {valid_ov(x.get("ov")) for x in candidates if valid_ov(x.get("ov"))}
                    new_ov = valid_ov(r.get("ov"))
                    same_ov = bool(new_ov and new_ov in candidate_ovs)
                    same_amount = any(
                        abs((r.get("cantidad") or 0) - (x.get("cantidad") or 0)) < 0.000001
                        and abs((r.get("pallets") or 0) - (x.get("pallets") or 0)) < 0.000001
                        for x in candidates
                    )
                    same_appointment = any(
                        daydiff(r.get("fecha_cita"), x.get("fecha_cita")) == 0
                        for x in candidates
                    )

                    if len(candidates) > 1 and not same_ov and not same_amount and not same_appointment:
                        result["nuevos"].append(r)
                        continue

                    result["ambiguos"].append({
                        "nuevo": r,
                        "candidatos": [{
                            "puntaje": similarity(r, x),
                            "proceso_id": x["proceso_id"],
                            "ov": x.get("ov"),
                            "sku": x.get("sku"),
                            "lote": x.get("lote"),
                            "finca": x.get("finca"),
                            "destino": x.get("destino"),
                            "cantidad": x.get("cantidad"),
                            "pallets": x.get("pallets"),
                            "fecha_proceso": x.get("fecha_proceso"),
                            "fecha_cita": x.get("fecha_cita"),
                            "ocurrencia": x.get("ocurrencia_consulta"),
                        } for x in sorted(candidates, key=lambda x: similarity(r, x), reverse=True)[:5]]
                    })
                    continue

        if not match:
            # Si no comparte la base productiva exacta con un registro existente,
            # se considera nuevo. Nunca se fusionan fincas/lotes/destinos distintos.
            result["nuevos"].append(r)
            continue

        used.add(match["proceso_id"])
        r["proceso_id"] = match["proceso_id"]
        r["metodo_coincidencia"] = method
        differences = business_changes(r, match)
        if not differences:
            result["sin_cambios"].append(r)
        else:
            result["modificados"].append({"nuevo": r, "anterior": match, "cambios": differences})

    weeks = {r["semana"] for r in records}
    for old in existing:
        if old.get("semana") in weeks and old["proceso_id"] not in used and old.get("activo_fuente", True):
            result["retirados"].append(old)
    return result

def fetch_existing(sb):
    rows=sb.table("procesos_fuente").select("*,procesos!inner(id,identidad_normalizada,activo_fuente)").limit(10000).execute().data
    out=[]
    for x in rows:
        p=x.pop("procesos");x["proceso_id"]=p["id"];x["identidad_normalizada"]=p["identidad_normalizada"];x["activo_fuente"]=p["activo_fuente"];out.append(x)
    return out

def source_payload(r,process_id,path):
    d={k:r.get(k) for k in SOURCE_FIELDS};d.update({"proceso_id":process_id,"archivo_origen":path.name,"datos_extra":{"fila_excel":r.get("fila_excel"),"ocurrencia_importador":r.get("ocurrencia_importador")}});return d

def main():
    ap=argparse.ArgumentParser();ap.add_argument("excel");ap.add_argument("--aplicar",action="store_true");ap.add_argument("--forzar-reimportacion",action="store_true",help="Permite reaplicar el mismo archivo para corregir precisión de datos");args=ap.parse_args();path=Path(args.excel)
    records,rejected=read_excel(path);url=os.getenv("SUPABASE_URL");key=os.getenv("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:raise SystemExit("Faltan SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY")
    from supabase import create_client
    sb=create_client(url,key);existing=fetch_existing(sb);changes=classify(records,existing)
    summary={"archivo":path.name,"filas_validas":len(records),"filas_rechazadas":len(rejected),"nuevos":len(changes["nuevos"]),"modificados":len(changes["modificados"]),"sin_cambios":len(changes["sin_cambios"]),"retirados":len(changes["retirados"]),"ambiguos":len(changes["ambiguos"]),"modo":"APLICAR" if args.aplicar else "VISTA_PREVIA"}
    print(json.dumps(summary,ensure_ascii=False,indent=2))
    if changes["ambiguos"]:
        print("\nAmbiguos detectados:")
        for item in changes["ambiguos"][:25]:
            nuevo=item["nuevo"]
            print(json.dumps({
                "fila_excel": nuevo.get("fila_excel"),
                "sku": nuevo.get("sku"),
                "lote": nuevo.get("lote"),
                "finca": nuevo.get("finca"),
                "destino": nuevo.get("destino"),
                "fecha_proceso": nuevo.get("fecha_proceso"),
                "fecha_cita": nuevo.get("fecha_cita"),
                "ocurrencia": nuevo.get("ocurrencia_consulta") or nuevo.get("ocurrencia_importador"),
                "candidatos": item.get("candidatos", []),
            }, ensure_ascii=False, indent=2))
    if changes["modificados"]:
        print("\nCambios de muestra:")
        for item in changes["modificados"][:10]:
            n=item["nuevo"];print(json.dumps({"proceso_id":n["proceso_id"],"metodo":n["metodo_coincidencia"],"cambios":item["cambios"]},ensure_ascii=False,indent=2))
    if not args.aplicar:
        print("\nVista previa únicamente. Para aplicar: agrega --aplicar")
        return
    if changes["ambiguos"]:
        raise SystemExit(f"BLOQUEADO: quedan {len(changes['ambiguos'])} ambiguos. No se escribio nada en Supabase.")
    fh=hash_file(path)
    if sb.table("importaciones").select("id").eq("hash_archivo",fh).execute().data and not args.forzar_reimportacion:raise SystemExit("Este archivo exacto ya fue importado. Si necesitas reparar precisión, usa --forzar-reimportacion.")
    imp=sb.table("importaciones").insert({"nombre_archivo":path.name,"hash_archivo":(fh if not args.forzar_reimportacion else fh+"-REPARACION-"+datetime.now(UTC).strftime("%Y%m%d%H%M%S")),"semanas":sorted({r["semana"] for r in records}),"filas_leidas":len(records)+len(rejected),"filas_validas":len(records),"filas_nuevas":len(changes["nuevos"]),"filas_modificadas":len(changes["modificados"]),"filas_retiradas":len(changes["retirados"]),"filas_ambiguas":len(changes["ambiguos"]),"estado":"VALIDADA"}).execute().data[0]
    for r in changes["nuevos"]:
        p=sb.table("procesos").insert({"identidad_normalizada":r["identidad_normalizada"],"activo_fuente":True,"estado_sincronizacion":"NUEVO","primera_importacion_id":imp["id"],"ultima_importacion_id":imp["id"]}).execute().data[0]
        sb.table("procesos_fuente").insert(source_payload(r,p["id"],path)).execute();sb.table("procesos_planificacion").insert({"proceso_id":p["id"]}).execute();sb.table("procesos_versiones").insert({"proceso_id":p["id"],"importacion_id":imp["id"],"version_numero":1,"tipo_cambio":"NUEVO","datos_nuevos":source_payload(r,p["id"],path)}).execute()
    for item in changes["modificados"]:
        r,o=item["nuevo"],item["anterior"];pid=r["proceso_id"]
        version=sb.table("procesos_versiones").select("version_numero").eq("proceso_id",pid).order("version_numero",desc=True).limit(1).execute().data;vn=(version[0]["version_numero"] if version else 0)+1
        newdata=source_payload(r,pid,path);sb.table("procesos_versiones").insert({"proceso_id":pid,"importacion_id":imp["id"],"version_numero":vn,"tipo_cambio":"MODIFICADO","datos_anteriores":o,"datos_nuevos":newdata}).execute();sb.table("procesos_fuente").update(newdata).eq("proceso_id",pid).execute();sb.table("procesos").update({"activo_fuente":True,"estado_sincronizacion":"MODIFICADO","ultima_importacion_id":imp["id"]}).eq("id",pid).execute()
    for r in changes["sin_cambios"]:sb.table("procesos").update({"activo_fuente":True,"estado_sincronizacion":"SIN_CAMBIOS","ultima_importacion_id":imp["id"]}).eq("id",r["proceso_id"]).execute()
    for o in changes["retirados"]:sb.table("procesos").update({"activo_fuente":False,"estado_sincronizacion":"RETIRADO_ORIGEN","ultima_importacion_id":imp["id"]}).eq("id",o["proceso_id"]).execute()
    for a in changes["ambiguos"]:sb.table("conciliaciones_pendientes").insert({"importacion_id":imp["id"],"registro_nuevo":a["nuevo"],"puntaje":a["candidatos"][0]["puntaje"] if a["candidatos"] else None}).execute()
    sb.table("importaciones").update({"estado":"APLICADA","aplicado_en":datetime.now(UTC).isoformat()}).eq("id",imp["id"]).execute();print("\nImportación incremental aplicada. La planificación manual no fue modificada.")
if __name__=="__main__":main()
