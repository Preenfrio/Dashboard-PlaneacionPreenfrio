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
def base_key(r):
    return "|".join([slug(r.get("ov") or "SIN-OV"),slug(r["sku"]),slug(r.get("lote") or "SIN-LOTE"),slug(r["finca"]),slug(r["destino"]),r["fecha_proceso"][:10]])
def identity(r):
    if norm(r.get("llave_embarque_consulta")):return "ORIGEN|"+norm(r["llave_embarque_consulta"])
    return f"{base_key(r)}|OC-{r['ocurrencia_importador']}"

def read_excel(path):
    df=pd.read_excel(path,sheet_name="PLANEACION",engine="openpyxl"); c=cols(df); out=[]; rejected=[]; seen={}; ovcol=c.get("OV") or c.get("OV OM")
    for i,row in df.iterrows():
        fp=dt(row[c["Fecha de proceso"]]); fc=dt(row[c["FECHA CITA"]]); pal=num(row[c["Pallets"]])
        reason="Pallets no numérico (por ejemplo, Granel)" if pal is None else "Pallets debe ser mayor que cero" if pal<=0 else "Fecha de proceso inválida" if not fp else "Fecha de cita inválida" if not fc else None
        if reason: rejected.append({"fila":i+2,"motivo":reason}); continue
        ov=None if not ovcol or pd.isna(row[ovcol]) else str(row[ovcol]).strip() or None
        r={"ov":ov,"descripcion":str(row[c["Descripción"]]).strip(),"sku":str(row[c["SKU"]]).strip(),"lote":None if pd.isna(row[c["Lote"]]) else str(row[c["Lote"]]).strip(),"finca":str(row[c["Finca"]]).strip(),"destino":str(row[c["Destino"]]).strip(),"cliente":None,"cantidad":num(row[c["Cantidad"]]),"pallets":pal,"fecha_proceso":fp,"fecha_cita":fc,"semana":integer(row[c["SEMANA"]]) if c.get("SEMANA") else pd.Timestamp(fp).isocalendar().week,"dias_transito":num(row[c["DIAS TRANSITO"]]) if c.get("DIAS TRANSITO") else None,"horas_transito":num(row[c["HORAS TRANSITO"]]) if c.get("HORAS TRANSITO") else None,"nivel_critico":integer(row[c["NIVEL CRITICO"]]) if c.get("NIVEL CRITICO") else None,"llave_base_consulta":None if not c.get("Llave_Base_Consulta") or pd.isna(row[c["Llave_Base_Consulta"]]) else str(row[c["Llave_Base_Consulta"]]).strip(),"ocurrencia_consulta":integer(row[c["Ocurrencia_Consulta"]]) if c.get("Ocurrencia_Consulta") else None,"llave_embarque_consulta":None if not c.get("Llave_Embarque_Consulta") or pd.isna(row[c["Llave_Embarque_Consulta"]]) else str(row[c["Llave_Embarque_Consulta"]]).strip(),"fila_excel":i+2}
        b=base_key(r);seen[b]=seen.get(b,0)+1;r["ocurrencia_importador"]=r["ocurrencia_consulta"] or seen[b];r["identidad_normalizada"]=identity(r)
        r["hash_contenido"]=hash_json({k:v for k,v in r.items() if k not in ("fila_excel","hash_contenido","identidad_normalizada")});out.append(r)
    return out,rejected

def daydiff(a,b):
    try:return abs((pd.Timestamp(a)-pd.Timestamp(b)).total_seconds())/86400
    except:return 9999

def similarity(new,old):
    score=0
    if norm(new.get("ov")) and norm(new.get("ov"))==norm(old.get("ov")):score+=55
    if norm(new["sku"])==norm(old.get("sku")):score+=25
    if norm(new.get("lote")) and norm(new.get("lote"))==norm(old.get("lote")):score+=30
    if norm(new["finca"])==norm(old.get("finca")):score+=12
    if norm(new["destino"])==norm(old.get("destino")):score+=12
    if daydiff(new["fecha_proceso"],old.get("fecha_proceso"))<=1:score+=10
    if abs((new.get("pallets") or 0)-(old.get("pallets") or 0))<0.001:score+=8
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


def classify(records,existing):
    by_identity={x["identidad_normalizada"]:x for x in existing}; by_source={norm(x.get("llave_embarque_consulta")):x for x in existing if norm(x.get("llave_embarque_consulta"))}
    used=set();result={"nuevos":[],"modificados":[],"sin_cambios":[],"ambiguos":[],"retirados":[]}
    for r in records:
        match=None;method=None
        if norm(r.get("llave_embarque_consulta")) in by_source:match=by_source[norm(r["llave_embarque_consulta"])];method="LLAVE_EMBARQUE"
        elif r["identidad_normalizada"] in by_identity:match=by_identity[r["identidad_normalizada"]];method="IDENTIDAD"
        else:
            candidates=[]
            for old in existing:
                if old["proceso_id"] in used:continue
                # Exige ancla fuerte: OV+SKU o SKU+lote.
                strong=(norm(r.get("ov")) and norm(r.get("ov"))==norm(old.get("ov")) and norm(r["sku"])==norm(old.get("sku"))) or (norm(r.get("lote")) and norm(r.get("lote"))==norm(old.get("lote")) and norm(r["sku"])==norm(old.get("sku")))
                if strong:candidates.append((similarity(r,old),old))
            candidates.sort(key=lambda x:x[0],reverse=True)
            if candidates and candidates[0][0]>=70 and (len(candidates)==1 or candidates[0][0]-candidates[1][0]>=15):match=candidates[0][1];method="CONCILIACION"
            elif candidates and candidates[0][0]>=70:
                result["ambiguos"].append({"nuevo":r,"candidatos":[{"puntaje":s,"proceso_id":o["proceso_id"],"finca":o.get("finca"),"destino":o.get("destino"),"pallets":o.get("pallets")} for s,o in candidates[:3]]});continue
        if not match:result["nuevos"].append(r);continue
        used.add(match["proceso_id"]);r["proceso_id"]=match["proceso_id"];r["metodo_coincidencia"]=method
        differences = business_changes(r, match)
        if not differences:
            result["sin_cambios"].append(r)
        else:
            result["modificados"].append({"nuevo": r, "anterior": match, "cambios": differences})
    weeks={r["semana"] for r in records}
    for old in existing:
        if old.get("semana") in weeks and old["proceso_id"] not in used and old.get("activo_fuente",True):result["retirados"].append(old)
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
    if changes["modificados"]:
        print("\nCambios de muestra:")
        for item in changes["modificados"][:10]:
            n=item["nuevo"];print(json.dumps({"proceso_id":n["proceso_id"],"metodo":n["metodo_coincidencia"],"cambios":item["cambios"]},ensure_ascii=False,indent=2))
    if not args.aplicar:
        print("\nVista previa únicamente. Para aplicar: agrega --aplicar")
        return
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
