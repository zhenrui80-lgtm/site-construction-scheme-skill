#!/usr/bin/env python3
"""Reject builder patterns that destroy template formatting or inherit project facts."""

from __future__ import annotations

import argparse
import ast
import json
import re
from pathlib import Path


WHOLE_TEXT_ATTRIBUTES = {"text"}
FORBIDDEN_NAMES = {"PREVIOUS_UTILS", "BASE_BUILDER", "OLD_PROJECT", "OLD_PROJECT_NAME"}
FORMAT_CLEANERS = re.compile(
    r"(?:calibrat|normaliz|clean)[_\- ]*(?:format|font|paragraph|spacing|indent)|"
    r"(?:format|font|paragraph|spacing|indent)[_\- ]*(?:calibrat|normaliz|clean)",
    re.IGNORECASE,
)


def dotted_name(node: ast.AST) -> str:
    parts: list[str] = []
    while isinstance(node, ast.Attribute):
        parts.append(node.attr)
        node = node.value
    if isinstance(node, ast.Name):
        parts.append(node.id)
    return ".".join(reversed(parts))


def audit_source(source: str, filename: str = "<builder>") -> dict:
    failures: list[str] = []
    warnings: list[str] = []
    run_text_assignments = 0
    try:
        tree = ast.parse(source, filename=filename)
    except SyntaxError as exc:
        return {
            "passed": False,
            "failures": [f"syntax error at line {exc.lineno}: {exc.msg}"],
            "warnings": [],
            "run_text_assignments": 0,
        }

    for node in ast.walk(tree):
        if isinstance(node, ast.Name) and node.id in FORBIDDEN_NAMES:
            failures.append(f"line {node.lineno}: forbidden cross-project name {node.id}")
        elif isinstance(node, (ast.Assign, ast.AnnAssign, ast.AugAssign)):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            for target in targets:
                if not isinstance(target, ast.Attribute) or target.attr not in WHOLE_TEXT_ATTRIBUTES:
                    continue
                owner = dotted_name(target.value).lower()
                if owner.endswith("run") or ".runs[" in owner:
                    run_text_assignments += 1
                elif owner.endswith("paragraph") or owner.endswith("cell") or owner in {"p", "para"}:
                    failures.append(
                        f"line {target.lineno}: whole-object text assignment `{dotted_name(target)}` destroys runs"
                    )
                else:
                    warnings.append(
                        f"line {target.lineno}: review text assignment `{dotted_name(target)}`; only run.text is safe"
                    )
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            if FORMAT_CLEANERS.search(node.name):
                failures.append(f"line {node.lineno}: forbidden format-cleaner definition `{node.name}`")
        elif isinstance(node, ast.Call):
            called = dotted_name(node.func)
            if FORMAT_CLEANERS.search(called):
                failures.append(f"line {node.lineno}: forbidden format-cleaner call `{called}`")
        elif isinstance(node, ast.Constant) and isinstance(node.value, str):
            value = node.value.replace("/", "\\")
            if re.search(r"_work_[^\\]+\\.*build_scheme\.py", value, re.IGNORECASE):
                failures.append(f"line {getattr(node, 'lineno', '?')}: references another project builder")

    if run_text_assignments == 0:
        warnings.append("no run.text assignment found; confirm replacements are run-level")
    return {
        "passed": not failures,
        "failures": sorted(set(failures)),
        "warnings": sorted(set(warnings)),
        "run_text_assignments": run_text_assignments,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("builder", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    builder = args.builder.expanduser().resolve()
    if not builder.is_file():
        raise SystemExit(f"builder not found: {builder}")
    result = {"builder": str(builder), **audit_source(builder.read_text(encoding="utf-8-sig"), str(builder))}
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(payload + "\n", encoding="utf-8")
    print(payload)
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
