#!/usr/bin/env python3
"""Actualizador idempotente de Finca para Planeador V2.1.
Puede ejecutarse aunque el servidor ya tenga Finca integrada.
Uso:
  py agregar_finca_v2_corregido.py servidor_dashboard_v2.py dashboard_preenfrio_v2.html
"""
from pathlib import Path
import re, shutil, sys

if len(sys.argv) != 3:
    raise SystemExit('Uso: py agregar_finca_v2_corregido.py servidor_dashboard_v2.py dashboard_preenfrio_v2.html')
server, dash = map(Path, sys.argv[1:])
for p in (server, dash):
    if not p.is_file():
        raise SystemExit(f'No existe el archivo: {p}')

# Respaldos únicos de esta ejecución.
shutil.copy2(server, server.with_name(server.name + '.antes_finca'))
shutil.copy2(dash, dash.with_name(dash.name + '.antes_finca'))

s = server.read_text(encoding='utf-8')
server_changed = False

# El servidor adjunto ya puede traer Finca. Solo se modifica si aún falta.
if "'Finca','Destino','Pallets'" not in s:
    replacements = [
        ("['Destino','Pallets','Fecha de proceso','FECHA CITA']", "['Finca','Destino','Pallets','Fecha de proceso','FECHA CITA']"),
        ("dest=str(r[c['Destino']]).strip();info=cat.get(norm(dest));", "finca=str(r[c['Finca']]).strip();dest=str(r[c['Destino']]).strip();info=cat.get(norm(dest));"),
        ("key=f'{norm(dest)}|{pdte.date()}|{appt.date()}|{info[\"client\"]}'", "key=f'{norm(finca)}|{norm(dest)}|{pdte.date()}|{appt.date()}|{info[\"client\"]}'"),
        ("'destination':dest,'client':info['client']", "'farm':finca,'destination':dest,'client':info['client']"),
        ("rf=['key','destination','client','process_date'", "rf=['key','farm','destination','client','process_date'"),
        ("'CLIENTE':x['client'],'DESTINO':x['destination']", "'CLIENTE':x['client'],'FINCA':x['farm'],'DESTINO':x['destination']"),
    ]
    for old, new in replacements:
        if old not in s:
            raise SystemExit(f'El servidor no tiene Finca y no pude reconocer este bloque: {old[:55]}')
        s = s.replace(old, new, 1)
    server.write_text(s, encoding='utf-8')
    server_changed = True

h = dash.read_text(encoding='utf-8')
html_changed = False

# CSS: agrega una columna de 175px después de Orden.
if 'grid-template-columns: 88px 72px 175px 120px' not in h and 'grid-template-columns:88px 72px 175px 120px' not in h:
    h2, n = re.subn(
        r'grid-template-columns:\s*88px\s+72px\s+120px\s+105px\s+120px\s+72px\s+105px\s+145px\s+145px\s+100px\s+minmax\(700px,\s*1fr\)',
        'grid-template-columns: 88px 72px 175px 120px 105px 120px 72px 105px 145px 145px 100px minmax(700px, 1fr)',
        h, count=1)
    if n != 1:
        raise SystemExit('No pude localizar la cuadrícula principal del dashboard.')
    h = h2; html_changed = True

# Encabezado dinámico dentro de JavaScript.
if '<span>Orden</span><span>Finca</span><span>Cliente</span>' not in h:
    h2, n = re.subn(
        r'<span>Fecha</span><span>Orden</span><span>Cliente</span>',
        '<span>Fecha</span><span>Orden</span><span>Finca</span><span>Cliente</span>',
        h, count=1)
    if n != 1:
        raise SystemExit('No pude localizar el encabezado Fecha / Orden / Cliente.')
    h = h2; html_changed = True

# Valor dinámico inmediatamente después de orderBtns.
if "<span>${x.farm||''}</span><span>${x.client}</span>" not in h:
    marker = '</span><span>${x.client}</span><span>${x.destination}</span>'
    if marker not in h:
        raise SystemExit('No pude localizar la fila dinámica para insertar Finca.')
    h = h.replace(marker, "</span><span>${x.farm||''}</span><span>${x.client}</span><span>${x.destination}</span>", 1)
    html_changed = True

# Tooltip.
if '<span>Finca:</span><b>${x.farm||\'\'}</b>' not in h:
    marker = '<div class="tip"><span>Cliente:</span><b>${x.client}</b></div>'
    if marker not in h:
        raise SystemExit('No pude localizar Cliente en el tooltip.')
    h = h.replace(marker, '<div class="tip"><span>Finca:</span><b>${x.farm||\'\'}</b></div>'+marker, 1)
    html_changed = True

if html_changed:
    dash.write_text(h, encoding='utf-8')

print('Actualización correcta.')
print('Servidor:', 'ya tenía Finca' if not server_changed else 'Finca agregada')
print('Dashboard:', 'ya tenía Finca' if not html_changed else 'Finca agregada después de Orden')
print('Respaldos creados con extensión .antes_finca')
