param([ValidateSet('Install','Remove')][string]$Action = 'Install')
# Auto-start for WalMartSummary when the PC starts or restarts, WITH auto-recovery (restarts the server if it stops).
$TaskName = 'WalMartSummary_BootRecovery'
$OtherTask = 'WalMartSummary_Boot'   # the "no recovery" boot task (22); only one starter may own the port
$Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$LoginLnk = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup\WalMartSummary.lnk'

if ($Action -eq 'Remove') {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Restart auto-start removed. (A server already running is not stopped; use '11 2 OFF'.)"
    } else { Write-Host 'Restart auto-start was not installed.' }
    return
}

$cfg = @{}
Get-Content (Join-Path $Root 'app_runtime.env') | ForEach-Object {
    if ($_ -match '^\s*([^#=\s]+)\s*=\s*(.*)$') { $cfg[$Matches[1].ToUpper()] = $Matches[2].Trim() }
}
$HostName = if ($cfg.HOST) { $cfg.HOST } else { '127.0.0.1' }
$Port = if ($cfg.PORT) { $cfg.PORT } else { '8040' }
$Py = if ($cfg.PYTHON_EXE -and (Test-Path $cfg.PYTHON_EXE)) { $cfg.PYTHON_EXE } else { (Get-Command python.exe -ErrorAction Stop).Source }
New-Item -ItemType Directory -Force (Join-Path $Root 'logs') | Out-Null

$cmd = "/c `"`"$Py`" -m uvicorn app.main:app --host $HostName --port $Port >> logs\server.log 2>&1`""
$taskAction = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument $cmd -WorkingDirectory $Root
$trigger = New-ScheduledTaskTrigger -AtStartup
$trigger.Delay = 'PT45S'   # let the network (server IP) come up first
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew `
    -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1)
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName $TaskName -Action $taskAction -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null

if (Get-ScheduledTask -TaskName $OtherTask -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $OtherTask -Confirm:$false
    Write-Host 'Removed the "no recovery" boot task (22) so only one starter exists.'
}
if (Test-Path $LoginLnk) { Remove-Item $LoginLnk; Write-Host 'Removed the login auto-start shortcut (33) so only one starter exists.' }
Write-Host "Done. WalMartSummary starts ~45s after the PC starts or restarts (port $Port) and is restarted if it stops."
Write-Host "To start it now without restarting: schtasks /Run /TN $TaskName"
