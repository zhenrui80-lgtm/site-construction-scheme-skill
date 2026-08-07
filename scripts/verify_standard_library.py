#!/usr/bin/env python3
import argparse
import hashlib
import json
from pathlib import Path

from pypdf import PdfReader


def sha256(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main():
    parser = argparse.ArgumentParser(description="Verify every indexed standard PDF and detect unindexed files.")
    parser.add_argument("--library-root", type=Path)
    args = parser.parse_args()

    skill_root = Path(__file__).resolve().parents[1]
    root = (args.library_root or (skill_root / "standards")).expanduser().resolve()
    index_path = root / "index.json"
    library = root / "library"
    data = json.loads(index_path.read_text(encoding="utf-8"))

    failures = []
    indexed_names = set()
    indexed_hashes = set()
    for record in data.get("standards", []):
        filename = record.get("filename")
        path = library / filename
        indexed_names.add(filename)
        if not path.exists():
            failures.append(f"missing file: {filename}")
            continue
        actual_hash = sha256(path)
        if actual_hash != record.get("sha256"):
            failures.append(f"hash mismatch: {filename}")
        if path.stat().st_size != record.get("size_bytes"):
            failures.append(f"size mismatch: {filename}")
        try:
            pages = len(PdfReader(str(path)).pages)
        except Exception as exc:
            failures.append(f"unreadable PDF: {filename}: {exc}")
            continue
        if pages != record.get("pages"):
            failures.append(f"page mismatch: {filename}: {pages} != {record.get('pages')}")
        if actual_hash in indexed_hashes:
            failures.append(f"duplicate indexed hash: {actual_hash}")
        indexed_hashes.add(actual_hash)

    unindexed = sorted(path.name for path in library.glob("*.pdf") if path.name not in indexed_names)
    failures.extend(f"unindexed PDF: {name}" for name in unindexed)
    result = {
        "library": str(library),
        "indexed": len(data.get("standards", [])),
        "pdf_files": len(list(library.glob("*.pdf"))),
        "failures": failures,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(1 if failures else 0)


if __name__ == "__main__":
    main()
