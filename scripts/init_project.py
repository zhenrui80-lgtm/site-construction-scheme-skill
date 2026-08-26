#!/usr/bin/env python3
"""Create a v2 work directory and validation contract without touching sources."""

from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path

from project_contract import write_contract


def main():
    parser = argparse.ArgumentParser(description="Initialize a site-construction-scheme v2 project contract.")
    parser.add_argument("workdir", type=Path)
    parser.add_argument("--project", required=True)
    template_group = parser.add_mutually_exclusive_group(required=True)
    template_group.add_argument("--template", type=Path)
    template_group.add_argument("--template-style", choices=("冷轧厂", "条材厂"))
    parser.add_argument("--profile", choices=("small", "standard", "complex"), default="standard")
    parser.add_argument("--contractor", default="武汉兴达建筑安装工程有限责任公司")
    parser.add_argument("--start-date")
    parser.add_argument("--finish-date")
    parser.add_argument("--duration-days", type=int)
    parser.add_argument("--basis-note")
    args = parser.parse_args()

    workdir = args.workdir.expanduser().resolve()
    if args.template:
        template = args.template.expanduser().resolve()
    else:
        template = Path(__file__).resolve().parent.parent / "assets" / f"{args.template_style}施工方案版式模板.docx"
    if not template.exists():
        raise SystemExit(f"Template not found: {template}")
    for name in ("intake", "assets", "output", "qa"):
        (workdir / name).mkdir(parents=True, exist_ok=True)
    template_copy = workdir / "intake" / template.name
    if template != template_copy:
        shutil.copy2(template, template_copy)

    if args.start_date or args.finish_date:
        if not args.start_date or not args.finish_date:
            raise SystemExit("Both --start-date and --finish-date are required for fixed dates")
        schedule = {
            "basis": "fixed_dates",
            "start_date": args.start_date,
            "finish_date": args.finish_date,
            "duration_days": args.duration_days,
        }
    else:
        if not args.duration_days:
            raise SystemExit("--duration-days is required when fixed dates are absent")
        schedule = {
            "basis": "day_index",
            "duration_days": args.duration_days,
            "basis_note": args.basis_note or "实际日历日期以建设单位开工或停机移交通知换算。",
        }

    contract = {
        "schema_version": 2,
        "project": {"name": args.project, "contractor": args.contractor},
        "document": {
            "template": str(template_copy.relative_to(workdir)),
            "builder": "build_scheme.py",
            "output_docx": f"output/{args.project.replace('#', '号')}施工方案.docx",
            "body_start": "一、编制依据及工程概况",
            "body_occurrence": 2,
            "body_font": "方正仿宋_GB2312",
            "profile": args.profile,
        },
        "schedule": schedule,
        "coverage": {
            "required_terms": [args.project],
            "forbidden_terms": [],
            "project_terms": [],
            "method_terms": [],
            "legacy_terms": [],
        },
        "qa": {
            "directory": "qa",
            "render_scale": 1.5,
            "allowed_sparse_pages": [1, 2, 3],
            "max_unexpected_sparse_pages": 0,
            "sparse_ink_ratio": 0.025,
            "expect_landscape_last_page": True,
        },
        "standards": {
            "verify_library": True,
            "strict_metadata": False,
            "required_files": [],
        },
    }
    output = write_contract(workdir / "project-contract.json", contract)
    print(json.dumps({"workdir": str(workdir), "contract": str(output), "template": str(template_copy)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
