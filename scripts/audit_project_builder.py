#!/usr/bin/env python3
"""Reject project builders that inherit code from a previous project script."""

from __future__ import annotations

import argparse
import ast
import json
import re
from pathlib import Path


FORBIDDEN_NAMES = {"PREVIOUS_UTILS", "BASE_BUILDER"}
WORK_SCRIPT = re.compile(r"(?:^|[\\/])work[\\/].*\.py$", re.IGNORECASE)


def audit_source(source, filename="<builder>"):
    failures = []
    try:
        tree = ast.parse(source, filename=filename)
    except SyntaxError as exc:
        return [f"builder syntax error at line {exc.lineno}: {exc.msg}"]

    for node in ast.walk(tree):
        if isinstance(node, ast.Name) and node.id in FORBIDDEN_NAMES:
            failures.append(f"line {node.lineno}: forbidden compatibility name {node.id}")
        elif isinstance(node, ast.Import):
            for alias in node.names:
                if alias.name == "build_scheme" or alias.name.endswith(".build_scheme"):
                    failures.append(f"line {node.lineno}: imports another build_scheme module")
        elif isinstance(node, ast.ImportFrom):
            module = node.module or ""
            if module == "build_scheme" or module.endswith(".build_scheme"):
                failures.append(f"line {node.lineno}: imports from another build_scheme module")
        elif isinstance(node, ast.Constant) and isinstance(node.value, str):
            value = node.value.strip()
            normalized = value.replace("/", "\\")
            if "build_scheme.py" in normalized.lower():
                failures.append(f"line {getattr(node, 'lineno', '?')}: references build_scheme.py")
            elif WORK_SCRIPT.search(value):
                failures.append(f"line {getattr(node, 'lineno', '?')}: references a Python script under work/")

    return sorted(set(failures))


def main():
    parser = argparse.ArgumentParser(description="Audit a project builder for cross-project code inheritance.")
    parser.add_argument("builder", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    builder = args.builder.expanduser().resolve()
    failures = audit_source(builder.read_text(encoding="utf-8-sig"), str(builder))
    result = {"builder": str(builder), "failures": failures, "passed": not failures}
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(payload + "\n", encoding="utf-8")
    print(payload)
    raise SystemExit(1 if failures else 0)


if __name__ == "__main__":
    main()
