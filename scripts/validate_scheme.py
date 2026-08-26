#!/usr/bin/env python3
"""Run the complete site-construction-scheme v2 delivery gate from one contract."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

from project_contract import load_contract, resolve_path, schedule_summary, validate_contract


ROOT = Path(__file__).resolve().parent


def run_command(label, command):
    completed = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return {
        "label": label,
        "command": [str(part) for part in command],
        "exit_code": completed.returncode,
        "stdout": completed.stdout[-4000:],
        "stderr": completed.stderr[-4000:],
    }


def append_terms(command, option, values):
    for value in values:
        command.extend((option, str(value)))


def main():
    parser = argparse.ArgumentParser(description="Run every v2 construction-scheme delivery gate.")
    parser.add_argument("contract", type=Path)
    parser.add_argument("--skip-word-export", action="store_true", help="Test-only: use qa.existing_pdf")
    args = parser.parse_args()

    contract_path, contract = load_contract(args.contract)
    contract_failures = validate_contract(contract_path, contract, require_files=True)
    qa_config = contract.get("qa", {})
    qa_dir = resolve_path(contract_path, qa_config.get("directory", "qa"))
    qa_dir.mkdir(parents=True, exist_ok=True)
    summary_path = qa_dir / "validation-summary.json"
    if contract_failures:
        result = {"contract": str(contract_path), "failures": contract_failures, "steps": []}
        summary_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        print(json.dumps(result, ensure_ascii=False, indent=2))
        raise SystemExit(1)

    document = contract["document"]
    coverage = contract.get("coverage", {})
    docx = resolve_path(contract_path, document["output_docx"])
    template = resolve_path(contract_path, document["template"])
    pdf = qa_dir / "final.pdf"
    steps = []
    failures = []

    builder = document.get("builder")
    if builder:
        steps.append(
            run_command(
                "builder-audit",
                [
                    sys.executable,
                    str(ROOT / "audit_project_builder.py"),
                    str(resolve_path(contract_path, builder)),
                    "--output",
                    str(qa_dir / "builder-audit.json"),
                ],
            )
        )

    if args.skip_word_export:
        existing_pdf = qa_config.get("existing_pdf")
        if not existing_pdf:
            failures.append("--skip-word-export requires qa.existing_pdf")
        else:
            pdf = resolve_path(contract_path, existing_pdf)
            if not pdf.exists():
                failures.append(f"qa.existing_pdf does not exist: {pdf}")
    else:
        powershell = shutil.which("pwsh") or shutil.which("powershell") or shutil.which("powershell.exe")
        if not powershell:
            failures.append("PowerShell not found for Word export")
        else:
            export = run_command(
                "word-export",
                [
                    powershell,
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(ROOT / "export_docx_word.ps1"),
                    "-InputDocx",
                    str(docx),
                    "-OutputPdf",
                    str(pdf),
                ],
            )
            steps.append(export)
            if export["exit_code"]:
                failures.append("Word field update/PDF export failed")
    if not pdf.exists():
        failures.append(f"final PDF not found: {pdf}")

    front_output = qa_dir / "front-matter-audit.json"
    front_command = [
        sys.executable,
        str(ROOT / "audit_front_matter.py"),
        str(template),
        str(docx),
        "--body-start",
        document["body_start"],
        "--occurrence",
        str(document.get("body_occurrence", 2)),
        "--output",
        str(front_output),
    ]
    for pattern in document.get("allowed_front_text_patterns", []):
        front_command.extend(("--allow-text-pattern", pattern))
    steps.append(run_command("front-matter-audit", front_command))

    docx_output = qa_dir / "docx-audit.json"
    docx_command = [
        sys.executable,
        str(ROOT / "audit_docx.py"),
        str(docx),
        "--body-start",
        document["body_start"],
        "--occurrence",
        str(document.get("body_occurrence", 2)),
        "--expected-body-font",
        document.get("body_font", "方正仿宋_GB2312"),
        "--output",
        str(docx_output),
    ]
    append_terms(docx_command, "--allowed-font", document.get("allowed_body_fonts", []))
    append_terms(docx_command, "--required", coverage.get("required_terms", []))
    append_terms(docx_command, "--forbidden", coverage.get("forbidden_terms", []))
    append_terms(docx_command, "--forbidden", coverage.get("legacy_terms", []))
    steps.append(run_command("docx-audit", docx_command))

    depth_output = qa_dir / "depth-audit.json"
    depth_command = [
        sys.executable,
        str(ROOT / "audit_scheme_depth.py"),
        str(docx),
        "--profile",
        document["profile"],
        "--output",
        str(depth_output),
    ]
    append_terms(depth_command, "--project-term", coverage.get("project_terms", []))
    append_terms(depth_command, "--method-term", coverage.get("method_terms", []))
    append_terms(depth_command, "--legacy-term", coverage.get("legacy_terms", []))
    steps.append(run_command("depth-audit", depth_command))

    schedule = schedule_summary(contract)
    schedule_output = qa_dir / "schedule-audit.json"
    schedule_command = [
        sys.executable,
        str(ROOT / "audit_schedule.py"),
        str(docx),
        "--basis",
        schedule["basis"],
        "--duration-days",
        str(schedule["duration_days"]),
        "--output",
        str(schedule_output),
    ]
    if schedule["basis"] == "fixed_dates":
        schedule_command.extend(("--start-date", schedule["start_date"], "--finish-date", schedule["finish_date"]))
    steps.append(run_command("schedule-audit", schedule_command))

    if pdf.exists():
        render_output = qa_dir / "render-audit.json"
        render_command = [
            sys.executable,
            str(ROOT / "render_pdf_qa.py"),
            str(pdf),
            "--output-dir",
            str(qa_dir / "render"),
            "--scale",
            str(qa_config.get("render_scale", 1.5)),
            "--max-unexpected-sparse-pages",
            str(qa_config.get("max_unexpected_sparse_pages", 0)),
            "--sparse-ink-ratio",
            str(qa_config.get("sparse_ink_ratio", 0.025)),
            "--blank-ink-ratio",
            str(qa_config.get("blank_ink_ratio", 0.001)),
            "--max-edge-ink-ratio",
            str(qa_config.get("max_edge_ink_ratio", 0.02)),
            "--output",
            str(render_output),
        ]
        append_terms(render_command, "--allowed-sparse-page", qa_config.get("allowed_sparse_pages", [1, 2, 3]))
        if qa_config.get("expect_landscape_last_page", True):
            render_command.append("--expect-landscape-last-page")
        steps.append(run_command("render-audit", render_command))

    standards = contract.get("standards", {})
    standards_root = resolve_path(
        contract_path,
        standards.get("library_root", str(ROOT.parent / "standards")),
    )
    if standards.get("verify_library", True):
        standards_command = [
            sys.executable,
            str(ROOT / "verify_standard_library.py"),
            "--library-root",
            str(standards_root),
        ]
        if standards.get("strict_metadata", False):
            standards_command.append("--strict-metadata")
        steps.append(
            run_command(
                "standards-library-audit",
                standards_command,
            )
        )
    for filename in standards.get("required_files", []):
        if not (standards_root / "library" / filename).exists():
            failures.append(f"required archived standard missing: {filename}")

    for step in steps:
        if step["exit_code"]:
            failures.append(f"{step['label']} failed")
    result = {
        "contract": str(contract_path),
        "project": contract["project"],
        "docx": str(docx),
        "pdf": str(pdf),
        "schedule": schedule,
        "qa_directory": str(qa_dir),
        "steps": steps,
        "failures": sorted(set(failures)),
        "passed": not failures,
    }
    summary_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({key: value for key, value in result.items() if key != "steps"}, ensure_ascii=False, indent=2))
    raise SystemExit(1 if failures else 0)


if __name__ == "__main__":
    main()
