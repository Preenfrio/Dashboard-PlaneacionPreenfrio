#!/usr/bin/env python3
from __future__ import annotations
import hashlib,json,os,subprocess,sys,time
from datetime import datetime
from pathlib import Path
BASE=Path(__file__).resolve().parent
EXCEL=Path(os.environ.get('PLANEACION_EXCEL_PATH', BASE/'PLANEACION_SEMANAL 1.xlsx'))
IMPORTER=BASE/'importador_incremental_supabase.py'
LOG=BASE/'sincronizacion_planeacion.log'
STATE=BASE/'.monitor_planeacion_estado.json'
POLL=int(os.environ.get('SYNC_POLL_SECONDS','10'))
DEBOUNCE=int(os.environ.get('SYNC_DEBOUNCE_SECONDS','180'))
MAX_RETIRED=float(os.environ.get('MAX_RETIRED_RATIO','0.20'))
def log(msg):
    line=f"{datetime.now().isoformat(timespec='seconds')} | {msg}"; print(line,flush=True)
    with LOG.open('a',encoding='utf-8') as f:f.write(line+'\n')
def sha(path):
    h=hashlib.sha256()
    with path.open('rb') as f:
        for c in iter(lambda:f.read(1048576),b''):h.update(c)
    return h.hexdigest()
def state():
    try:return json.loads(STATE.read_text(encoding='utf-8'))
    except:return {}
def save(x):STATE.write_text(json.dumps(x,ensure_ascii=False,indent=2),encoding='utf-8')
def summary(output):
    i=output.find('{')
    if i<0:raise RuntimeError('El importador no devolvio JSON')
    return json.JSONDecoder().raw_decode(output[i:])[0]
def run(apply=False):
    cmd=[sys.executable,str(IMPORTER),str(EXCEL)]+(['--aplicar'] if apply else [])
    p=subprocess.run(cmd,cwd=BASE,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=600,env=os.environ.copy())
    out=(p.stdout or '')+'\n'+(p.stderr or '')
    if p.returncode:raise RuntimeError(out.strip())
    return summary(out)
def readable():
    try:
        import openpyxl
        w=openpyxl.load_workbook(EXCEL,read_only=True,data_only=True)
        ok={'PLANEACION','CATALOGO','UBICACIONES'}.issubset(w.sheetnames);w.close();return ok
    except Exception as e:log(f'Archivo todavia no disponible: {e}');return False
def stable():
    old=None;equal=0
    while equal<3:
        sig=(EXCEL.stat().st_size,EXCEL.stat().st_mtime_ns)
        equal=equal+1 if sig==old else 0;old=sig;time.sleep(POLL)
    return readable()
def synchronize():
    digest=sha(EXCEL); s=state()
    if s.get('ultimo_hash_aplicado')==digest:return log('Mismo archivo; no se importa')
    pre=run(False);log('Vista previa: '+json.dumps(pre,ensure_ascii=False))
    valid=int(pre.get('filas_validas',0));amb=int(pre.get('ambiguos',0));ret=int(pre.get('retirados',0));mod=int(pre.get('modificados',0));same=int(pre.get('sin_cambios',0));new=int(pre.get('nuevos',0));total=ret+mod+same
    ratio=ret/total if total else 0
    if valid<=0:return log('BLOQUEADO: sin filas validas')
    if amb>0:return log(f'BLOQUEADO: {amb} ambiguos')
    if ratio>MAX_RETIRED:return log(f'BLOQUEADO: retirados {ratio:.2%}')
    if new+mod+ret==0:save({'ultimo_hash_aplicado':digest});return log('Sin cambios de negocio')
    app=run(True);save({'ultimo_hash_aplicado':digest,'ultima_importacion':datetime.now().isoformat(),'resumen':pre});log('APLICADO: '+json.dumps(app,ensure_ascii=False))
def main():
    miss=[x for x in ('SUPABASE_URL','SUPABASE_SERVICE_ROLE_KEY') if not os.environ.get(x)]
    if miss:raise RuntimeError('Faltan variables: '+', '.join(miss))
    if not EXCEL.is_file() or not IMPORTER.is_file():raise RuntimeError('Falta Excel o importador en la carpeta')
    log(f'Monitor iniciado: {EXCEL}')
    last=(EXCEL.stat().st_size,EXCEL.stat().st_mtime_ns);pending=time.monotonic()-DEBOUNCE
    while True:
        try:
            sig=(EXCEL.stat().st_size,EXCEL.stat().st_mtime_ns)
            if sig!=last:last=sig;pending=time.monotonic();log('Cambio detectado; esperando 3 minutos')
            if pending is not None and time.monotonic()-pending>=DEBOUNCE:
                if stable():synchronize();last=(EXCEL.stat().st_size,EXCEL.stat().st_mtime_ns);pending=None
                else:pending=time.monotonic()
        except Exception as e:log('ERROR: '+str(e));pending=time.monotonic()
        time.sleep(POLL)
if __name__=='__main__':main()
