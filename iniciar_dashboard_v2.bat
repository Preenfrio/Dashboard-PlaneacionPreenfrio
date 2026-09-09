@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (py servidor_dashboard_v2.py) else (python servidor_dashboard_v2.py)
pause
