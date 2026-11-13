param([int]$Port = 8040, [switch]$DryRun)

# Stop only this app: the uvicorn "app.main:app" process listening on $Port (and its reload worker).
# Anything else on the machine, even if it is Python, is left running.
$listenerIds = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty OwningProcess -Unique
if (-not $listenerIds) { Write-Host "Nothing is listening on port $Port."; exit 0 }

foreach ($listenerId in $listenerIds) {
    $target = $null
    $currentId = $listenerId
    for ($depth = 0; $depth -lt 4 -and $currentId; $depth++) {
        $proc = Get-CimInstance Win32_Process -Filter "ProcessId=$currentId" -ErrorAction SilentlyContinue
        if (-not $proc) { break }
        if ($proc.CommandLine -match 'app\.main:app') { $target = $proc; break }
        $currentId = $proc.ParentProcessId
    }
    if (-not $target) {
        Write-Host "Port $Port is used by PID $listenerId, which is not WalMartSummary. Left running."
        continue
    }
    if ($DryRun) { Write-Host "Would stop PID $($target.ProcessId): $($target.CommandLine)"; continue }
    Write-Host "Stopping WalMartSummary PID $($target.ProcessId) (and its workers)."
    & taskkill /PID $target.ProcessId /T /F | Out-Null
}
