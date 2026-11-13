"""WalMartSummary background launcher: runs the app hidden, shows a tray icon,
restarts the server if it crashes, and shows Windows notifications."""
import ctypes
import os
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path

import pystray
from PIL import Image, ImageDraw

BASE = Path(__file__).resolve().parents[1]  # project root (folder with app_runtime.env)


def read_env() -> dict:
    values = {}
    env_file = BASE / "app_runtime.env"
    if env_file.exists():
        for line in env_file.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, value = line.split("=", 1)
                values[key.strip().upper()] = value.strip()
    return values


ENV = read_env()
HOST = ENV.get("HOST") or "127.0.0.1"
CHECK_HOST = "127.0.0.1" if HOST == "0.0.0.0" else HOST
PORT = int(ENV.get("PORT") or 8040)
URL = f"http://{CHECK_HOST}:{PORT}"
LOG_DIR = BASE / "logs"
LOG_FILE = LOG_DIR / "server.log"
ATTACH = "--attach" in sys.argv  # tray only: watch a server started elsewhere, never start/restart it
COLORS = {"running": "#1a9a4a", "starting": "#e0a100", "down": "#c62828"}


def already_running() -> bool:
    """Only one tray app at a time (named Windows mutex)."""
    global _mutex
    _mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "WalMartSummaryTray")
    return ctypes.windll.kernel32.GetLastError() == 183  # ERROR_ALREADY_EXISTS


def make_icon(state: str) -> Image.Image:
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((4, 4, 60, 60), fill=COLORS[state])
    d.text((22, 24), "WM", fill="white")
    return img


def port_open() -> bool:
    try:
        with socket.create_connection((CHECK_HOST, PORT), timeout=0.5):
            return True
    except OSError:
        return False


class Supervisor:
    def __init__(self):
        self.icon = None
        self.proc = None
        self.stopping = threading.Event()
        self.manual_restart = False

    def set_state(self, state: str, text: str, notify: str | None = None):
        if self.icon:
            self.icon.icon = make_icon(state)
            self.icon.title = f"WalMartSummary - {text}"
            self.icon.visible = state == "running"  # tray icon only while the app URL is active
            if notify:
                try:
                    self.icon.notify(notify, "WalMartSummary")
                except Exception:
                    pass

    def spawn(self):
        LOG_DIR.mkdir(exist_ok=True)
        log = open(LOG_FILE, "ab")
        cmd = [sys.executable, "-m", "uvicorn", "app.main:app", "--host", HOST, "--port", str(PORT)]
        return subprocess.Popen(cmd, cwd=BASE, stdout=log, stderr=subprocess.STDOUT,
                                stdin=subprocess.DEVNULL, creationflags=subprocess.CREATE_NO_WINDOW)

    def watch(self):
        last = None
        while not self.stopping.is_set():
            up = port_open()
            if up != last:
                last = up
                if up:
                    self.set_state("running", f"running at {URL}")
                else:
                    self.set_state("down", "server not running")
            self.stopping.wait(2)

    def run(self):
        if ATTACH:
            return self.watch()
        delay = 3
        while not self.stopping.is_set():
            started = time.time()
            self.set_state("starting", "starting...")
            self.proc = self.spawn()
            ready = False
            while self.proc.poll() is None and not self.stopping.is_set():
                if not ready and port_open():
                    ready = True
                    self.set_state("running", f"running at {URL}", f"WalMartSummary is running at {URL}")
                time.sleep(1)
            if self.stopping.is_set():
                break
            if self.manual_restart:
                self.manual_restart = False
                continue
            self.set_state("down", "stopped - restarting", "WalMartSummary stopped unexpectedly. Restarting...")
            delay = 3 if time.time() - started > 60 else min(delay * 2, 60)
            self.stopping.wait(delay)

    def restart(self, *_):
        if self.proc and self.proc.poll() is None:
            self.manual_restart = True
            self.proc.terminate()

    def stop_app(self, icon, *_):
        """Stop the app: end the server, then the tray icon (also when this tray only attached to a server)."""
        if ATTACH:
            subprocess.Popen(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(Path(__file__).with_name("stop_background.ps1"))],
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
        self.quit(icon)

    def quit(self, icon, *_):
        self.stopping.set()
        if self.proc and self.proc.poll() is None:
            self.proc.terminate()
        icon.stop()


def main():
    if already_running():
        return
    sup = Supervisor()
    menu = pystray.Menu(
        pystray.MenuItem("Open app", lambda *_: webbrowser.open(URL), default=True),
        *([] if ATTACH else [pystray.MenuItem("Restart server", sup.restart)]),
        pystray.MenuItem("Open log folder", lambda *_: (LOG_DIR.mkdir(exist_ok=True), os.startfile(LOG_DIR))),
        pystray.MenuItem("Stop the app", sup.stop_app),
    )
    sup.icon = pystray.Icon("WalMartSummary", make_icon("starting"), "WalMartSummary - starting...", menu)
    threading.Thread(target=sup.run, daemon=True).start()
    sup.icon.run(setup=lambda icon: setattr(icon, "visible", port_open()))


if __name__ == "__main__":
    main()
