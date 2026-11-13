"""In-memory progress and cancellation for long-running processes.

The UI polls /progress/status?token=... and can call /progress/cancel to stop everything.
"""
import subprocess
import threading
import time

_lock = threading.Lock()
_jobs: dict[str, dict] = {}
_processes: dict[str, subprocess.Popen] = {}
_KEEP_SECONDS = 3600
_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


class JobCancelled(Exception):
    """Raised inside a running job when the user pressed Stop."""


def update(token: str, percent: float, stage: str, **extra) -> None:
    if not token:
        return
    now = time.time()
    with _lock:
        job = _jobs.setdefault(token, {"started": now, "done": False})
        if job.get("cancelled") and not extra.get("done"):
            raise JobCancelled("Cancelled by user")
        job.update(percent=100 if extra.get("done") else max(job.get("percent", 0), percent), stage=stage, updated=now)
        job.update(extra)
        for key in [key for key, value in _jobs.items() if now - value.get("updated", now) > _KEEP_SECONDS]:
            del _jobs[key]


def bump(token: str, stage: str) -> None:
    """Advance a little and show the latest step text (used by Logger for flows without fixed stages)."""
    if not token:
        return
    with _lock:
        job = _jobs.get(token)
        current = job.get("percent", 10) if job else 10
    update(token, min(90, current + max(1, (90 - current) * 0.12)), stage[:160])


def get(token: str) -> dict | None:
    with _lock:
        job = _jobs.get(token)
        return dict(job) if job else None


def is_cancelled(token: str) -> bool:
    with _lock:
        return bool(token and _jobs.get(token, {}).get("cancelled"))


def register_process(token: str, process: subprocess.Popen) -> None:
    if token:
        with _lock:
            _processes[token] = process


def release_process(token: str) -> None:
    with _lock:
        _processes.pop(token, None)


def _kill_tree(pid: int) -> None:
    subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True, creationflags=_NO_WINDOW, check=False)


def _kill_automation_excel(since_epoch: float) -> None:
    """Close hidden Excel instances this app started (never the user's own open workbooks)."""
    script = (
        f"$since = [DateTimeOffset]::FromUnixTimeSeconds({int(since_epoch) - 5}).LocalDateTime; "
        "Get-CimInstance Win32_Process -Filter \"Name='EXCEL.EXE'\" | "
        "Where-Object { $_.CommandLine -match '/automation' -and $_.CreationDate -ge $since } | "
        "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
    )
    subprocess.run(
        ["powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive", "-Command", script],
        capture_output=True,
        creationflags=_NO_WINDOW,
        timeout=20,
        check=False,
    )


def cancel_jobs(token: str = "") -> int:
    """Cancel every running job (plus this token), kill their Excel/PowerShell work, and return how many were stopped."""
    now = time.time()
    with _lock:
        targets = [key for key, job in _jobs.items() if not job.get("done")]
        if token and token not in targets:
            _jobs.setdefault(token, {"started": now, "done": False})
            targets.append(token)
        earliest = now
        for key in targets:
            job = _jobs[key]
            job["cancelled"] = True
            job["stage"] = "Cancelled by user"
            earliest = min(earliest, job.get("started", now))
        processes = [_processes[key] for key in targets if key in _processes]
    for process in processes:
        try:
            _kill_tree(process.pid)
        except Exception:
            pass
    if processes or any(_jobs.get(key, {}).get("excel") for key in targets):
        try:
            _kill_automation_excel(earliest)
        except Exception:
            pass
    return len(targets)
