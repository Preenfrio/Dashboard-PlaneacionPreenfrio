MONITOR AUTOMATICO SHAREPOINT/ONEDRIVE -> SUPABASE

Coloca estos archivos junto a PLANEACION_SEMANAL 1.xlsx e importador_incremental_supabase.py.

PRUEBA MANUAL
1. Abre PowerShell en la carpeta.
2. Configura SUPABASE_URL y SUPABASE_SERVICE_ROLE_KEY en esa ventana.
3. Ejecuta: py monitor_planeacion_sharepoint.py
4. Guarda un cambio en el Excel compartido.
5. Revisa sincronizacion_planeacion.log.

INSTALACION AUTOMATICA
powershell -ExecutionPolicy Bypass -File .\configurar_monitor_windows.ps1

El monitor espera 3 minutos, comprueba estabilidad, bloquea ambiguos y bloquea retirados superiores al 20%.
No subas el log, el estado ni secretos a GitHub.
