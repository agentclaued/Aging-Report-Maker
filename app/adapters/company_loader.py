from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[2]
COMPANIES_DIR = BASE_DIR / "companies"


@dataclass(frozen=True)
class CompanyProfile:
    code: str
    display_name: str
    sql_database: str
    report_workflow_ready: bool
    steps: dict[str, dict]


def _read_json(path: Path) -> dict:
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def load_company_profile(company_code: str) -> CompanyProfile:
    normalized_code = company_code.strip().lower()
    company_dir = COMPANIES_DIR / normalized_code
    company_meta = _read_json(company_dir / "company.json")
    if not company_meta:
        raise FileNotFoundError(f"Company config not found for '{normalized_code}'")

    steps = {
        "step_01": _read_json(company_dir / "step_01.json"),
        "step_02": _read_json(company_dir / "step_02.json"),
        "step_03": _read_json(company_dir / "step_03.json"),
    }
    return CompanyProfile(
        code=str(company_meta.get("code", normalized_code)),
        display_name=str(company_meta.get("display_name", normalized_code.title())),
        sql_database=str(company_meta.get("sql_database", f"{normalized_code}DB")),
        report_workflow_ready=bool(company_meta.get("report_workflow_ready", False)),
        steps=steps,
    )
