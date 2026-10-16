#!/usr/bin/env python3
"""Read-only prose hints and required-text checks. Never rewrites a deliverable."""
import argparse
import hashlib
import json
import re
from pathlib import Path
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

PHRASES = ("仅供参考", "不承担任何责任", "无法保证准确", "作为一个AI", "作为人工智能", "不给审稿人递刀子")

def read_text(path):
    if path.suffix.lower() in (".md", ".txt"):
        return path.read_text(encoding="utf-8-sig")
    if path.suffix.lower() == ".docx":
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        return "\n".join("".join(n.text or "" for n in p.findall(".//w:t", ns)) for p in root.findall(".//w:p", ns))
    raise ValueError("Supported inputs: UTF-8 .md/.txt and .docx; export workbook narrative as text first.")

def normalize(s):
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", s))

def review(text, required=()):
    hints = [{"line": i, "phrase": p, "text": line.strip(), "action": "人工结合语境判断；保留必要事实和限制"}
             for i, line in enumerate(text.splitlines(), 1) for p in PHRASES if p in line]
    missing = [s for s in required if normalize(s) not in normalize(text)]
    return {"status": "missing_required_text" if missing else "review_hints" if hints else "no_lexical_hints",
            "hints": hints, "missing_required_text": missing,
            "scope": "只读词语提示与指定文本检查；不是技术正确性、完整性或可执行性批准"}

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--required-text", action="append", default=[])
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        if args.output and args.output.resolve() == args.input.resolve():
            raise ValueError("Output must differ from source; source is never overwritten.")
        source_hash = hashlib.sha256(args.input.read_bytes()).hexdigest()
        result = review(read_text(args.input), args.required_text)
        result["source_sha256"] = source_hash
        payload = json.dumps(result, ensure_ascii=False, indent=2)
        if args.output:
            args.output.write_text(payload + "\n", encoding="utf8")
        else:
            print(payload)
        return 1 if result["missing_required_text"] else 0
    except (OSError, ValueError, zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
        print(str(exc), file=sys.stderr)
        return 2

if __name__ == "__main__":
    raise SystemExit(main())
