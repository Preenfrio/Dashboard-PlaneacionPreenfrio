#!/usr/bin/env python3
from __future__ import annotations
import argparse, hashlib, json, os, re
from collections import Counter
from datetime import UTC, datetime
from pathlib import Path
import pandas as pd

REQUIRED=["Descripción","SKU","Lote","Finca","Destino","Cantidad","Pallets","Fecha de proceso","FECHA CITA"]
OPTIONAL=["SEMANA","DIAS TRANSITO","HORAS TRANSITO","NIVEL CRITICO","OV","OV OM","Llave_Base_Consulta","Ocurrencia_Consulta","Llave_Embarque_Consulta"]
SOURCE_FIELDS=["llave_base_consulta","ocurrencia_consulta","llave_embarque_consulta","ov","descripcion","sku","lote","finca","destino","cliente","cantidad","pallets","fecha_proceso","semana","fecha_cita","dias_transito","horas_transito","nivel_critico","hash_contenido"]

def norm(v): return " ".join(str(v if v is not None else "").strip().upper().split())
def slug(v): return re.sub(r"[^A-Z0-9]+","-",norm(v)).strip("-")
def dt(v):
    x=pd.to_datetime(v,dayfirst=True,errors="coerce"); return None if pd.isna(x) else x.to_pydatetime().isoformat(timespec="seconds")
def num(v):
    try:return None if pd.isna(v) else float(v)
    except:return None
def integer(v):
    try:return None if pd.isna(v) else int(float(v))
    except:return None
def hash_json(v): return hashlib.sha256(json.dumps(v,ensure_ascii=False,sort_keys=True,default=str).encode()).hexdigest()
def hash_file(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def cols(df):
    m={norm(c):c for c in df.columns}; missing=[x for x in REQUIRED if norm(x) not in m]
    if missing: raise ValueError("Faltan columnas: "+", ".join(missing))
    return {x:m.get(norm(x)) for x in REQUIRED+OPTIONAL}
def valid_ov(v):
    s=norm(v)
    if not s or s in {"0","NAN","NONE"}: return None
    return None if any(x in s for x in ("OV ND","SIN OV","NO COMPATIBLE","NO ENCONTRADA","NO ENCONTRADO","PENDIENTE","ERROR")) else s

def canonical_base(r):
    fp=(r.get("fecha_proceso") or "")[:10].replace("-",""); fc=(r.get("fecha_cita") or "")[:10].replace("-","")
    return "|".join([slug(r.get("finca")),fp,fc,slug(r.get("destino")),slug(r.get("sku")),slug(r.get("lote"))])
def canonical_key(r):
    occ=r.get("ocurrencia_importador") or r.get("ocurrencia_consulta") or 1
    return "CANON|"+canonical_base(r)+"|"+str(int(occ))
def strict_base_key(r):
    return "|".join([norm(r.get("sku")),norm(r.get("lote")),norm(r.get("finca")),norm(r.get("destino")),(r.get("fecha_proceso") or "")[:10]])
def daydiff(a,b):
    try:return abs((pd.Timestamp(a)-pd.Timestamp(b)).total_seconds())/86400
    except:return 9999

def read_excel(path):
    df=pd.read_excel(path,sheet_name="PLANEACION",engine="openpyxl"); c=cols(df); out=[]; rejected=[]; seen=Counter(); ovcol=c.get("OV") or c.get("OV OM")
    for i,row in df.iterrows():
        fp=dt(row[c["Fecha de proceso"]]); fc=dt(row[c["FECHA CITA"]]); pal=num(row[c["Pallets"]])
        reason="Pallets no numerico (por ejemplo, Granel)" if pal is None else "Pallets debe ser mayor que cero" if pal<=0 else "Fecha de proceso invalida" if not fp else "Fecha de cita invalida" if not fc else None
        if reason: rejected.append({"fila":i+2,"motivo":reason}); continue
        sem=integer(row[c["SEMANA"]]) if c.get("SEMANA") else None
        r={"ov":None if not ovcol or pd.isna(row[ovcol]) else valid_ov(row[ovcol]),"descripcion":str(row[c["Descripción"]]).strip(),"sku":str(row[c["SKU"]]).strip(),"lote":None if pd.isna(row[c["Lote"]]) else str(row[c["Lote"]]).strip(),"finca":str(row[c["Finca"]]).strip(),"destino":str(row[c["Destino"]]).strip(),"cliente":None,"cantidad":num(row[c["Cantidad"]]),"pallets":pal,"fecha_proceso":fp,"fecha_cita":fc,"semana":sem if sem is not None else int(pd.Timestamp(fp).isocalendar().week),"dias_transito":num(row[c["DIAS TRANSITO"]]) if c.get("DIAS TRANSITO") else None,"horas_transito":num(row[c["HORAS TRANSITO"]]) if c.get("HORAS TRANSITO") else None,"nivel_critico":integer(row[c["NIVEL CRITICO"]]) if c.get("NIVEL CRITICO") else None,"llave_base_consulta":None if not c.get("Llave_Base_Consulta") or pd.isna(row[c["Llave_Base_Consulta"]]) else str(row[c["Llave_Base_Consulta"]]).strip(),"ocurrencia_consulta":integer(row[c["Ocurrencia_Consulta"]]) if c.get("Ocurrencia_Consulta") else None,"llave_embarque_consulta":None if not c.get("Llave_Embarque_Consulta") or pd.isna(row[c["Llave_Embarque_Consulta"]]) else str(row[c["Llave_Embarque_Consulta"]]).strip(),"fila_excel":i+2}
        base=canonical_base(r); seen[base]+=1; r["ocurrencia_importador"]=seen[base]; r["identidad_normalizada"]=canonical_key(r)
        r["hash_contenido"]=hash_json({k:v for k,v in r.items() if k not in {"fila_excel","hash_contenido","identidad_normalizada"}}); out.append(r)
    grouped={}
    for r in out: grouped.setdefault(r["identidad_normalizada"],[]).append(r["fila_excel"])
    dup={k:v for k,v in grouped.items() if len(v)>1}
    if dup: raise ValueError("Identidades internas duplicadas: "+json.dumps(dict(list(dup.items())[:10]),ensure_ascii=False))
    return out,rejected

def comparable(field,v):
    if field in {"cantidad","pallets","dias_transito","horas_transito"}: return None if v is None else round(float(v),6)
    if field in {"semana","nivel_critico","ocurrencia_consulta"}: return None if v is None else int(v)
    if field in {"fecha_proceso","fecha_cita"}:
        x=pd.to_datetime(v,errors="coerce"); return None if pd.isna(x) else x.to_pydatetime().isoformat(timespec="seconds")
    return None if v is None else norm(v)
def business_changes(n,o):
    fs=["llave_base_consulta","ocurrencia_consulta","llave_embarque_consulta","ov","descripcion","sku","lote","finca","destino","cliente","cantidad","pallets","fecha_proceso","semana","fecha_cita","dias_transito","horas_transito","nivel_critico"]
    return {f:{"antes":o.get(f),"despues":n.get(f)} for f in fs if comparable(f,o.get(f))!=comparable(f,n.get(f))}
def fetch_existing(sb):
    rows=sb.table("procesos_fuente").select("*,procesos!inner(id,identidad_normalizada,activo_fuente)").limit(10000).execute().data; out=[]
    for original in rows:
        x=dict(original); p=x.pop("procesos"); x.update(proceso_id=p["id"],identidad_normalizada=p["identidad_normalizada"],activo_fuente=p["activo_fuente"]); out.append(x)
    return out

def classify(records,existing):
    """Clasifica registros protegiendo la identidad de origen.

    Precedencia:
    1. Llave_Embarque_Consulta exacta, cuando viene informada.
       - Una coincidencia exacta conserva el proceso.
       - Sin coincidencia exacta se clasifica como NUEVO.
       - Nunca baja a conciliacion aproximada contra otras ocurrencias.
    2. Sin llave de embarque: identidad CANON exacta.
    3. Sin llave de embarque: conciliacion por base productiva estricta.
    """
    byid={}; bysource={}; bybase={}
    for x in existing:
        byid.setdefault(x.get("identidad_normalizada"),[]).append(x)
        source_key=norm(x.get("llave_embarque_consulta"))
        if source_key:
            bysource.setdefault(source_key,[]).append(x)
        bybase.setdefault(strict_base_key(x),[]).append(x)

    used=set()
    result={"nuevos":[],"modificados":[],"sin_cambios":[],"ambiguos":[],"retirados":[]}

    for r in records:
        match=None; method=None
        source_key=norm(r.get("llave_embarque_consulta"))

        # Una llave de embarque valida es la identidad autoritativa.
        # Si no existe exactamente en Supabase, es una partida nueva;
        # no debe compararse contra otras ocurrencias parecidas.
        if source_key:
            all_source_matches=bysource.get(source_key,[])
            available_source_matches=[
                x for x in all_source_matches
                if x["proceso_id"] not in used
            ]

            if len(available_source_matches)==1:
                match=available_source_matches[0]
                method="LLAVE_EMBARQUE"
            elif len(available_source_matches)>1:
                result["ambiguos"].append({
                    "nuevo":r,
                    "motivo":"LLAVE_EMBARQUE_DUPLICADA_EN_SUPABASE",
                    "candidatos":available_source_matches[:5],
                })
                continue
            elif all_source_matches:
                result["ambiguos"].append({
                    "nuevo":r,
                    "motivo":"LLAVE_EMBARQUE_REPETIDA_EN_ARCHIVO_ACTUAL",
                    "candidatos":all_source_matches[:5],
                })
                continue
            else:
                result["nuevos"].append(r)
                continue

        # Solo los registros sin llave de embarque pueden usar CANON
        # y la conciliacion aproximada por base productiva.
        if not match:
            identity_matches=[
                x for x in byid.get(r["identidad_normalizada"],[])
                if x["proceso_id"] not in used
            ]
            if len(identity_matches)==1:
                match=identity_matches[0]
                method="IDENTIDAD"
            elif len(identity_matches)>1:
                result["ambiguos"].append({
                    "nuevo":r,
                    "motivo":"IDENTIDAD_CANON_DUPLICADA",
                    "candidatos":identity_matches[:5],
                })
                continue

        if not match:
            candidates=[
                x for x in bybase.get(strict_base_key(r),[])
                if x["proceso_id"] not in used
            ]
            if candidates:
                ov=valid_ov(r.get("ov"))
                tests=[
                    (lambda x:bool(ov and valid_ov(x.get("ov"))==ov),"BASE+OV"),
                    (lambda x:abs((r.get("cantidad") or 0)-(x.get("cantidad") or 0))<1e-6 and abs((r.get("pallets") or 0)-(x.get("pallets") or 0))<1e-6,"BASE+CANTIDAD"),
                    (lambda x:daydiff(r.get("fecha_cita"),x.get("fecha_cita"))==0,"BASE+CITA"),
                ]
                for test,name in tests:
                    rows=[x for x in candidates if test(x)]
                    if len(rows)==1:
                        match,method=rows[0],name
                        break
                if not match and len(candidates)==1:
                    match,method=candidates[0],"BASE_UNICA"
                if not match:
                    result["ambiguos"].append({"nuevo":r,"candidatos":candidates[:5]})
                    continue

        if not match:
            result["nuevos"].append(r)
            continue

        used.add(match["proceso_id"])
        r.update(proceso_id=match["proceso_id"],metodo_coincidencia=method)
        diff=business_changes(r,match)
        result["modificados" if diff else "sin_cambios"].append(
            {"nuevo":r,"anterior":match,"cambios":diff} if diff else r
        )

    weeks={r["semana"] for r in records}
    result["retirados"]=[
        x for x in existing
        if x.get("semana") in weeks
        and x["proceso_id"] not in used
        and x.get("activo_fuente",True)
    ]
    return result


def source_payload(r,pid,path):
    d={k:r.get(k) for k in SOURCE_FIELDS}; d.update(proceso_id=pid,archivo_origen=path.name,datos_extra={"fila_excel":r.get("fila_excel"),"ocurrencia_importador":r.get("ocurrencia_importador")}); return d

def get_or_create_process(sb,r,impid):
    rows=sb.table("procesos").select("id").eq("identidad_normalizada",r["identidad_normalizada"]).limit(1).execute().data
    if rows:
        pid=rows[0]["id"]; sb.table("procesos").update({"activo_fuente":True,"estado_sincronizacion":"NUEVO","ultima_importacion_id":impid}).eq("id",pid).execute(); return pid
    return sb.table("procesos").insert({"identidad_normalizada":r["identidad_normalizada"],"activo_fuente":True,"estado_sincronizacion":"NUEVO","primera_importacion_id":impid,"ultima_importacion_id":impid}).execute().data[0]["id"]

def ensure_version(sb,pid,impid,tipo_cambio,datos_nuevos,datos_anteriores=None):
    """Crea una sola versión por proceso/importación/tipo y calcula el número siguiente.

    Hace idempotente la recuperación de importaciones parciales: si la misma
    versión ya se escribió antes del error, no vuelve a insertarla.
    """
    existing=(sb.table("procesos_versiones").select("proceso_id,version_numero")
        .eq("proceso_id",pid).eq("importacion_id",impid)
        .eq("tipo_cambio",tipo_cambio).limit(1).execute().data)
    if existing:
        return existing[0]["version_numero"]
    versions=(sb.table("procesos_versiones").select("version_numero")
        .eq("proceso_id",pid).order("version_numero",desc=True)
        .limit(1).execute().data)
    version_number=(int(versions[0]["version_numero"]) if versions else 0)+1
    record={"proceso_id":pid,"importacion_id":impid,"version_numero":version_number,
            "tipo_cambio":tipo_cambio,"datos_nuevos":datos_nuevos}
    if datos_anteriores is not None:
        record["datos_anteriores"]=datos_anteriores
    sb.table("procesos_versiones").insert(record).execute()
    return version_number

def ensure_children(sb,r,pid,impid,path):
    payload=source_payload(r,pid,path)
    if sb.table("procesos_fuente").select("proceso_id").eq("proceso_id",pid).limit(1).execute().data:
        sb.table("procesos_fuente").update(payload).eq("proceso_id",pid).execute()
    else:
        sb.table("procesos_fuente").insert(payload).execute()
    if not sb.table("procesos_planificacion").select("proceso_id").eq("proceso_id",pid).limit(1).execute().data:
        sb.table("procesos_planificacion").insert({"proceso_id":pid}).execute()
    ensure_version(sb,pid,impid,"NUEVO",payload)

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("excel"); ap.add_argument("--aplicar",action="store_true"); ap.add_argument("--forzar-reimportacion",action="store_true"); args=ap.parse_args(); path=Path(args.excel)
    records,rejected=read_excel(path); url=os.getenv("SUPABASE_URL"); key=os.getenv("SUPABASE_SERVICE_ROLE_KEY") or os.getenv("SUPABASE_SECRET_KEY")
    if not url or not key: raise SystemExit("Faltan SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY")
    from supabase import create_client
    sb=create_client(url,key); changes=classify(records,fetch_existing(sb)); summary={"archivo":path.name,"filas_validas":len(records),"filas_rechazadas":len(rejected),**{k:len(changes[k]) for k in ("nuevos","modificados","sin_cambios","retirados","ambiguos")},"modo":"APLICAR" if args.aplicar else "VISTA_PREVIA"}; print(json.dumps(summary,ensure_ascii=False,indent=2))
    if not args.aplicar: print("\nVista previa unicamente. Para aplicar: agrega --aplicar"); return
    if changes["ambiguos"]: raise SystemExit(f"BLOQUEADO: quedan {len(changes['ambiguos'])} ambiguos. No se escribio nada.")
    fh=hash_file(path); rows=sb.table("importaciones").select("*").eq("hash_archivo",fh).limit(1).execute().data; previous=rows[0] if rows else None
    if previous and previous.get("estado")=="APLICADA" and not args.forzar_reimportacion: raise SystemExit("Este archivo exacto ya fue importado correctamente.")
    payload={"nombre_archivo":path.name,"hash_archivo":fh if not args.forzar_reimportacion else fh+"-REPARACION-"+datetime.now(UTC).strftime("%Y%m%d%H%M%S"),"semanas":sorted({int(r["semana"]) for r in records}),"filas_leidas":len(records)+len(rejected),"filas_validas":len(records),"filas_nuevas":len(changes["nuevos"]),"filas_modificadas":len(changes["modificados"]),"filas_retiradas":len(changes["retirados"]),"filas_ambiguas":0,"estado":"VALIDADA"}
    if previous and previous.get("estado")!="APLICADA" and not args.forzar_reimportacion: sb.table("importaciones").update(payload).eq("id",previous["id"]).execute(); imp={**previous,**payload}; print("\nRecuperando importacion incompleta anterior.")
    else: imp=sb.table("importaciones").insert(payload).execute().data[0]
    try:
        for r in changes["nuevos"]: pid=get_or_create_process(sb,r,imp["id"]); ensure_children(sb,r,pid,imp["id"],path)
        for item in changes["modificados"]:
            r,o=item["nuevo"],item["anterior"]
            pid=r["proceso_id"]
            data=source_payload(r,pid,path)
            ensure_version(sb,pid,imp["id"],"MODIFICADO",data,o)
            sb.table("procesos_fuente").update(data).eq("proceso_id",pid).execute()
            sb.table("procesos").update({"activo_fuente":True,"estado_sincronizacion":"MODIFICADO","ultima_importacion_id":imp["id"]}).eq("id",pid).execute()
        for r in changes["sin_cambios"]: sb.table("procesos").update({"activo_fuente":True,"estado_sincronizacion":"SIN_CAMBIOS","ultima_importacion_id":imp["id"]}).eq("id",r["proceso_id"]).execute()
        for o in changes["retirados"]: sb.table("procesos").update({"activo_fuente":False,"estado_sincronizacion":"RETIRADO_ORIGEN","ultima_importacion_id":imp["id"]}).eq("id",o["proceso_id"]).execute()
        sb.table("importaciones").update({"estado":"APLICADA","aplicado_en":datetime.now(UTC).isoformat()}).eq("id",imp["id"]).execute()
    except Exception:
        sb.table("importaciones").update({"estado":"ERROR"}).eq("id",imp["id"]).execute(); raise
    print("\nImportacion incremental aplicada. La planificacion manual no fue modificada.")
if __name__=="__main__": main()
