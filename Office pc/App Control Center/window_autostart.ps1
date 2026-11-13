param([ValidateSet('Install','Remove')][string]$Action = 'Install')
# Runs WalMartSummary like the other office tools: a boot task that opens a visible console window (its own venv).
$TaskName = 'Tools-walmart_summary'
$Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).ProviderPath  # plain path (works for UNC shares too)

if ($Action -eq 'Remove') {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Removed task $TaskName. (A running window is not closed; close it yourself.)"
    } else { Write-Host 'Task was not installed.' }
    return
}

$cfg = @{}
Get-Content (Join-Path $Root 'app_runtime.env') | ForEach-Object {
    if ($_ -match '^\s*([^#=\s]+)\s*=\s*(.*)$') { $cfg[$Matches[1].ToUpper()] = $Matches[2].Trim() }
}
$HostName = if ($cfg.HOST) { $cfg.HOST } else { '0.0.0.0' }
$Port = if ($cfg.PORT) { $cfg.PORT } else { '8507' }

# Use the Python from PYTHON_EXE in app_runtime.env when it exists (e.g. a venv kept on C:); otherwise fall back to <app folder>\venv.
$VenvPy = if ($cfg.PYTHON_EXE -and (Test-Path -LiteralPath $cfg.PYTHON_EXE)) { $cfg.PYTHON_EXE } else { Join-Path $Root 'venv\Scripts\python.exe' }
if (-not (Test-Path -LiteralPath $VenvPy)) {
    Write-Host 'Creating the Python environment (venv)...'
    & python -m venv (Join-Path $Root 'venv')
    if ($LASTEXITCODE -ne 0) { throw 'Could not create venv. Is Python installed and on PATH?' }
}
Write-Host 'Installing packages...'
& $VenvPy -m pip install -r (Join-Path $Root 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'pip install failed.' }

$user = "$env:USERDOMAIN\$env:USERNAME"
$cred = Get-Credential -UserName $user -Message 'Windows password of the user that will run WalMartSummary (needed so it can start at boot)'
# One console window; if the server stops, uvicorn is started again after 5 seconds (loop inside the window).
$cmd = "/k title WalMartSummary $Port && pushd `"$Root`" && for /l %i in (1,1,999999) do (`"$VenvPy`" -m uvicorn app.main:app --host $HostName --port $Port & timeout /t 5 /nobreak >nul)"
$taskAction = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument $cmd
# Runs at PC start/restart, at every login, and re-checks every 10 minutes (a second copy is ignored while the window is open),
# so closing the window or a crash is repaired automatically.
$bootTrigger = New-ScheduledTaskTrigger -AtStartup
$bootTrigger.Delay = 'PT45S'
$logonTrigger = New-ScheduledTaskTrigger -AtLogOn
$repeatTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 10) -RepetitionDuration (New-TimeSpan -Days 3650)
$trigger = @($bootTrigger, $logonTrigger, $repeatTrigger)
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
    -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName $TaskName -Action $taskAction -Trigger $trigger -Settings $settings `
    -User $cred.UserName -Password $cred.GetNetworkCredential().Password -RunLevel Highest -Force | Out-Null
Write-Host "Done. WalMartSummary starts ~45s after the PC starts or restarts, at every login, and is restarted within 10 minutes if it stops (port $Port)."
Write-Host "Start it now: schtasks /Run /TN $TaskName"
