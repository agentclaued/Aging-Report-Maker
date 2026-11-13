param([ValidateSet('Install','Remove')][string]$Action = 'Install')
# Boot auto-start for WalMartSummary: Task Scheduler task at PC startup (before any login).
# Starts the server once, hidden, NO restart on failure (no auto-recovery).
$TaskName = 'WalMartSummary_Boot'
$Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$TrayLnk = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup\WalMartSummary (tray).lnk'
$LoginLnk = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup\WalMartSummary.lnk'

if ($Action -eq 'Remove') {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        if (Test-Path $TrayLnk) { Remove-Item $TrayLnk }
        Write-Host "Boot auto-start removed. (A server already running is not stopped; use '11 2 OFF'.)"
    } else { Write-Host 'Boot auto-start was not installed.' }
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
    -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew -RestartCount 0
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName $TaskName -Action $taskAction -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null

# The login tray app would also supervise/restart the same port, so remove it.
if (Test-Path $LoginLnk) { Remove-Item $LoginLnk; Write-Host 'Removed the login auto-start shortcut (tray, auto-recovery) so only one starter exists.' }
# tray-only icon at login (attaches to the boot server; never starts or restarts it)
$Pyw = Join-Path (Split-Path $Py) 'pythonw.exe'
$sc = (New-Object -ComObject WScript.Shell).CreateShortcut($TrayLnk)
$sc.TargetPath = $Pyw; $sc.Arguments = '"' + (Join-Path $PSScriptRoot 'background_tray.pyw') + '" --attach'
$sc.WorkingDirectory = $Root; $sc.WindowStyle = 7; $sc.Save()
Write-Host "Done. WalMartSummary starts ~45s after the PC boots (port $Port), with no auto-recovery."
Write-Host "To start it now without rebooting: schtasks /Run /TN $TaskName"
