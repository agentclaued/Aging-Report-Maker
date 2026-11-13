# Stops only THIS project's tray app and its server. Other apps (even other Python or uvicorn apps) are left alone.
$controlDir = $PSScriptRoot
$projectRoot = (Resolve-Path (Join-Path $controlDir '..')).Path

$port = 8040
$envFile = Join-Path $projectRoot 'app_runtime.env'
if (Test-Path $envFile) {
    foreach ($line in Get-Content $envFile) {
        if ($line -match '^\s*PORT\s*=\s*(\d+)') { $port = [int]$Matches[1] }
    }
}

# 1) This folder's tray app (matched by its full path) together with the server it started.
Get-CimInstance Win32_Process |
    Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -like "*$controlDir*background_tray.pyw*" } |
    ForEach-Object { & taskkill /PID $_.ProcessId /T /F | Out-Null }

# 2) A leftover server on this app's port, only if it is the app.main:app uvicorn (itself or its parent).
Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty OwningProcess -Unique |
    ForEach-Object {
        $currentId = $_
        for ($depth = 0; $depth -lt 4 -and $currentId; $depth++) {
            $proc = Get-CimInstance Win32_Process -Filter "ProcessId=$currentId" -ErrorAction SilentlyContinue
            if (-not $proc) { break }
            if ($proc.CommandLine -match 'app\.main:app') { & taskkill /PID $proc.ProcessId /T /F | Out-Null; break }
            $currentId = $proc.ParentProcessId
        }
    }
