$ErrorActionPreference="Stop"
$base=Split-Path -Parent $MyInvocation.MyCommand.Path
$excel=Join-Path $base "PLANEACION_SEMANAL 1.xlsx"
$monitor=Join-Path $base "monitor_planeacion_sharepoint.py"
if(-not(Test-Path $excel)){throw "No se encontro $excel"}
$url=Read-Host "Pega SUPABASE_URL"
$secret=Read-Host "Pega SUPABASE_SERVICE_ROLE_KEY" -AsSecureString
$plain=[System.Net.NetworkCredential]::new("",$secret).Password
[Environment]::SetEnvironmentVariable("SUPABASE_URL",$url,"User")
[Environment]::SetEnvironmentVariable("SUPABASE_SERVICE_ROLE_KEY",$plain,"User")
[Environment]::SetEnvironmentVariable("PLANEACION_EXCEL_PATH",$excel,"User")
[Environment]::SetEnvironmentVariable("MAX_RETIRED_RATIO","0.20","User")
$py=(Get-Command py.exe).Source
$action=New-ScheduledTaskAction -Execute $py -Argument ('"'+$monitor+'"') -WorkingDirectory $base
$trigger=New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$settings=New-ScheduledTaskSettingsSet -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
$principal=New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName "Monitor Planeacion SharePoint Supabase" -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force|Out-Null
Write-Host "Monitor instalado. Inicia la tarea desde Programador de tareas o cierra y abre sesion."
