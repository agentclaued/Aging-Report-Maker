from __future__ import annotations

import ast
from datetime import datetime
from pathlib import Path
import re


BASE_DIR = Path(__file__).resolve().parent.parent
LOG_DIR = BASE_DIR / "edit log"
LOG_FILE = LOG_DIR / "ai edit history.txt"
APP_RUN_LOG_FILE = LOG_DIR / "app run logs.txt"
APP_DIR = BASE_DIR / "app"

INTRO_LINES = [
    "AI EDIT HISTORY",
    "Project: overStockap",
    f"Location: {BASE_DIR}",
    "Rule: Append only. Do not delete old history lines. Add new updates below old ones with date and time.",
    "",
]


def _python_files() -> list[Path]:
    return sorted(path for path in APP_DIR.rglob("*.py") if path.name != "__init__.py")


def _collect_function_index() -> list[str]:
    lines = [
        "FUNCTION INDEX SNAPSHOT",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "Purpose: Quick file and line reference for edits and runtime error checks.",
        "",
    ]

    for path in _python_files():
        rel_path = path.relative_to(BASE_DIR)
        # Accept UTF-8 files with a BOM so startup logging does not fail on parse.
        tree = ast.parse(path.read_text(encoding="utf-8-sig"), filename=str(path))
        file_lines = [str(rel_path)]

        for node in tree.body:
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                file_lines.append(f"- line {node.lineno}: {node.name}()")
            elif isinstance(node, ast.ClassDef):
                for child in node.body:
                    if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef)):
                        file_lines.append(f"- line {child.lineno}: {node.name}.{child.name}()")

        if len(file_lines) > 1:
            lines.extend(file_lines)
            lines.append("")

    return lines


def _read_existing_history() -> str:
    if not LOG_FILE.exists():
        return ""

    content = LOG_FILE.read_text(encoding="utf-8")
    marker = "\nEDIT HISTORY\n"
    if marker in content:
        return content.rsplit(marker, 1)[1].lstrip()
    return content.strip()


def _next_letter(history_body: str, date_text: str) -> str:
    current_date = None
    letters: list[str] = []

    for raw_line in history_body.splitlines():
        line = raw_line.strip()
        date_match = re.match(r"^\[(\d{4}-\d{2}-\d{2})\]$", line)
        if date_match:
            current_date = date_match.group(1)
            continue

        if current_date != date_text:
            continue

        letter_match = re.match(r"^([A-Z])\.", line)
        if letter_match:
            letters.append(letter_match.group(1))

    if not letters:
        return "A"

    return chr(ord(max(letters)) + 1)


def append_history_event(title: str, details: list[str]) -> Path:
    LOG_DIR.mkdir(parents=True, exist_ok=True)

    history_body = _read_existing_history()
    date_text = datetime.now().strftime("%Y-%m-%d")
    time_text = datetime.now().strftime("%H:%M:%S")
    letter = _next_letter(history_body, date_text)

    entry_lines: list[str] = []
    if f"[{date_text}]" not in history_body:
        if history_body and not history_body.endswith("\n"):
            history_body += "\n"
        if history_body:
            history_body += "\n"
        entry_lines.append(f"[{date_text}]")

    entry_lines.append(f"{letter}. {title}")
    entry_lines.append(f"- Time: {time_text}")
    entry_lines.extend(f"- {detail}" for detail in details)

    appended_block = "\n".join(entry_lines).strip()
    if history_body:
        if not history_body.endswith("\n"):
            history_body += "\n"
        history_body += "\n" + appended_block + "\n"
    else:
        history_body = appended_block + "\n"

    final_lines = INTRO_LINES + _collect_function_index() + ["EDIT HISTORY", "", history_body.rstrip(), ""]
    LOG_FILE.write_text("\n".join(final_lines), encoding="utf-8")
    return LOG_FILE


def append_app_run_log(title: str, details: list[str]) -> Path:
    LOG_DIR.mkdir(parents=True, exist_ok=True)

    now = datetime.now()
    date_text = now.strftime("%Y-%m-%d")
    time_text = now.strftime("%H:%M:%S")

    entry_lines = [
        f"[{date_text}]",
        title,
        f"- Time: {time_text}",
        *[f"- {detail}" for detail in details],
        "",
    ]

    with APP_RUN_LOG_FILE.open("a", encoding="utf-8") as handle:
        handle.write("\n".join(entry_lines))

    return APP_RUN_LOG_FILE
