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
    parser.add_argument("--strict-metadata", action="store_true")
    args = parser.parse_args()

    skill_root = Path(__file__).resolve().parents[1]
    root = (args.library_root or (skill_root / "standards")).expanduser().resolve()
    index_path = root / "index.json"
    library = root / "library"
    data = json.loads(index_path.read_text(encoding="utf-8"))

    failures = []
    metadata_warnings = []
    indexed_names = set()
    indexed_hashes = set()
    for record in data.get("standards", []):
        filename = record.get("filename")
        if not record.get("number") or record.get("number") == "UNKNOWN":
            metadata_warnings.append(f"missing standard number: {filename}")
        if not record.get("title"):
            metadata_warnings.append(f"missing title: {filename}")
        sources = record.get("sources", [])
        if not sources:
            metadata_warnings.append(f"missing source metadata: {filename}")
        for source in sources:
            if source.get("url") == "local-file" and not source.get("kind"):
                metadata_warnings.append(f"legacy local source without kind: {filename}")
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
    if args.strict_metadata:
        failures.extend(metadata_warnings)
    result = {
        "library": str(library),
        "indexed": len(data.get("standards", [])),
        "pdf_files": len(list(library.glob("*.pdf"))),
        "metadata_warnings": sorted(set(metadata_warnings)),
        "failures": failures,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(1 if failures else 0)


if __name__ == "__main__":
    main()
