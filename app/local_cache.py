"""Temporary SQLite cache (app/local.db) so a workbook is parsed once, not on every request."""
import hashlib
import pickle
import sqlite3
import time
from pathlib import Path
from typing import Callable

import pandas as pd

CACHE_DB = Path(__file__).resolve().parent / "local.db"
MAX_ENTRIES = 8
MAX_AGE_SECONDS = 24 * 60 * 60


def _connect() -> sqlite3.Connection:
    connection = sqlite3.connect(CACHE_DB, timeout=10)
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")
    connection.execute("CREATE TABLE IF NOT EXISTS frame_cache (key TEXT PRIMARY KEY, created REAL, payload BLOB)")
    return connection


def _file_key(file_path: str, label: str) -> str:
    digest = hashlib.sha1()
    with open(file_path, "rb") as file_handle:
        for chunk in iter(lambda: file_handle.read(1 << 20), b""):
            digest.update(chunk)
    return f"{label}:{digest.hexdigest()}"


def cached_frame(file_path: str, label: str, loader: Callable[[], pd.DataFrame]) -> pd.DataFrame:
    """Return the DataFrame for this file content from local.db, loading and storing it on a miss."""
    try:
        key = _file_key(file_path, label)
        with _connect() as connection:
            row = connection.execute("SELECT payload FROM frame_cache WHERE key = ?", (key,)).fetchone()
            if row:
                return pickle.loads(row[0])
    except Exception:
        key = ""

    frame = loader()

    if key:
        try:
            with _connect() as connection:
                connection.execute(
                    "INSERT OR REPLACE INTO frame_cache (key, created, payload) VALUES (?, ?, ?)",
                    (key, time.time(), pickle.dumps(frame, protocol=pickle.HIGHEST_PROTOCOL)),
                )
                connection.execute("DELETE FROM frame_cache WHERE created < ?", (time.time() - MAX_AGE_SECONDS,))
                connection.execute(
                    "DELETE FROM frame_cache WHERE key NOT IN (SELECT key FROM frame_cache ORDER BY created DESC LIMIT ?)",
                    (MAX_ENTRIES,),
                )
        except Exception:
            pass
    return frame
