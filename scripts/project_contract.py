#!/usr/bin/env python3
"""Load and validate the v2 project contract shared by build and QA scripts."""

from __future__ import annotations

import json
from datetime import date
from pathlib import Path


SCHEMA_VERSION = 2
PROFILES = {"small", "standard", "complex"}
SCHEDULE_BASES = {"fixed_dates", "day_index"}


def load_contract(path):
    contract_path = Path(path).expanduser().resolve()
    data = json.loads(contract_path.read_text(encoding="utf-8"))
    return contract_path, data


def resolve_path(contract_path, value):
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = contract_path.parent / path
    return path.resolve()


def parse_iso_date(value, field):
    try:
        return date.fromisoformat(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field} must be YYYY-MM-DD: {value!r}") from exc


def schedule_summary(contract):
    schedule = contract.get("schedule", {})
    basis = schedule.get("basis")
    if basis not in SCHEDULE_BASES:
        raise ValueError(f"schedule.basis must be one of {sorted(SCHEDULE_BASES)}")
    if basis == "fixed_dates":
        start = parse_iso_date(schedule.get("start_date"), "schedule.start_date")
        finish = parse_iso_date(schedule.get("finish_date"), "schedule.finish_date")
        if finish < start:
            raise ValueError("schedule.finish_date precedes schedule.start_date")
        computed = (finish - start).days + 1
        supplied = schedule.get("duration_days")
        if supplied is not None and int(supplied) != computed:
            raise ValueError(f"schedule.duration_days {supplied} != inclusive duration {computed}")
        return {"basis": basis, "start_date": start.isoformat(), "finish_date": finish.isoformat(), "duration_days": computed}
    duration = int(schedule.get("duration_days", 0))
    if duration < 1:
        raise ValueError("schedule.duration_days must be positive for day_index basis")
    note = str(schedule.get("basis_note", "")).strip()
    if not note:
        raise ValueError("schedule.basis_note is required for day_index basis")
    return {"basis": basis, "duration_days": duration, "basis_note": note}


def validate_contract(contract_path, contract, require_files=True):
    failures = []
    if contract.get("schema_version") != SCHEMA_VERSION:
        failures.append(f"schema_version must be {SCHEMA_VERSION}")
    project = contract.get("project", {})
    if not str(project.get("name", "")).strip():
        failures.append("project.name is required")
    document = contract.get("document", {})
    if document.get("profile") not in PROFILES:
        failures.append(f"document.profile must be one of {sorted(PROFILES)}")
    for field in ("template", "output_docx", "body_start"):
        if not str(document.get(field, "")).strip():
            failures.append(f"document.{field} is required")
    if int(document.get("body_occurrence", 0)) < 1:
        failures.append("document.body_occurrence must be positive")
    coverage = contract.get("coverage", {})
    profile = document.get("profile")
    minimums = {"small": (4, 5), "standard": (6, 8), "complex": (8, 12)}
    if profile in minimums:
        project_min, method_min = minimums[profile]
        if len(coverage.get("project_terms", [])) < project_min:
            failures.append(f"coverage.project_terms requires at least {project_min} entries for {profile}")
        if len(coverage.get("method_terms", [])) < method_min:
            failures.append(f"coverage.method_terms requires at least {method_min} entries for {profile}")
    try:
        schedule_summary(contract)
    except ValueError as exc:
        failures.append(str(exc))
    if require_files:
        for field in ("template", "output_docx", "builder"):
            value = document.get(field)
            if value and not resolve_path(contract_path, value).exists():
                failures.append(f"document.{field} does not exist: {resolve_path(contract_path, value)}")
    return failures


def write_contract(path, contract):
    output = Path(path).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(contract, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return output
