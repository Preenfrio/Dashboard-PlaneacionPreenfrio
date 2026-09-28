#!/usr/bin/env python3
from pathlib import Path
import py_compile
import shutil

p = Path('importador_incremental_supabase.py')
if not p.exists():
    raise SystemExit('ERROR: no encuentro importador_incremental_supabase.py')

text = p.read_text(encoding='utf-8')
backup = p.with_suffix(p.suffix + '.antes_semana_v7')
shutil.copy2(p, backup)

old_field = '"semana":integer(row[c["SEMANA"]]) if c.get("SEMANA") else pd.Timestamp(fp).isocalendar().week'
new_field = '"semana":(integer(row[c["SEMANA"]]) if c.get("SEMANA") and integer(row[c["SEMANA"]]) is not None else int(pd.Timestamp(fp).isocalendar().week))'

if old_field not in text:
    raise SystemExit('ERROR: no encontré la asignación de semana esperada; no modifiqué nada')
text = text.replace(old_field, new_field, 1)

old_sorted = '"semanas":sorted({r["semana"] for r in records})'
new_sorted = '"semanas":sorted({int(r["semana"]) for r in records if r.get("semana") is not None})'
if old_sorted not in text:
    raise SystemExit('ERROR: no encontré el sorted de semanas esperado; no modifiqué nada')
text = text.replace(old_sorted, new_sorted, 1)

# Aserciones de seguridad antes de guardar
compile(text, str(p), 'exec')
assert new_field in text
assert new_sorted in text

p.write_text(text, encoding='utf-8')
py_compile.compile(str(p), doraise=True)

print('OK: semana vacía ahora usa la semana ISO de Fecha de proceso.')
print('OK: el registro de importación ya no intenta ordenar int junto con None.')
print('OK: importador_incremental_supabase.py compiló correctamente.')
print(f'Respaldo: {backup.name}')
