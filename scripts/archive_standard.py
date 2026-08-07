#!/usr/bin/env python3
import argparse
import hashlib
import json
import re
import shutil
from datetime import date, datetime, timezone
from pathlib import Path

from pypdf import PdfReader


def sha256(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def safe_filename(name):
    cleaned = re.sub(r"[<>:\"/\\|?*]+", "_", name).strip(" .")
    return cleaned or "standard.pdf"


def infer_number(stem):
    match = re.search(r"(?i)(GBT|GB/T|GB|JGJT|JGJ/T|JGJ|SHT|SH/T|SH|TSG|DLT|DL/T|DL|YBT|YB/T|YB)\s*([A-Z]?\d+(?:\.\d+)?)\s*[-—_]\s*(\d{4})", stem)
    if not match:
        return "UNKNOWN"
    prefix = match.group(1).upper().replace(" ", "")
    prefix = {
        "GBT": "GB/T",
        "JGJT": "JGJ/T",
        "SHT": "SH/T",
        "DLT": "DL/T",
        "YBT": "YB/T",
    }.get(prefix, prefix)
    return f"{prefix} {match.group(2)}-{match.group(3)}"


def load_index(path):
    if not path.exists():
        return {"schema_version": 1, "updated_at": None, "standards": []}
    return json.loads(path.read_text(encoding="utf-8"))


def write_index(path, data):
    data["updated_at"] = datetime.now(timezone.utc).isoformat()
    data["standards"] = sorted(
        data["standards"],
        key=lambda item: (item.get("number", ""), item.get("title", ""), item.get("sha256", "")),
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def append_unique(items, value):
    if value and value not in items:
        items.append(value)


def main():
    parser = argparse.ArgumentParser(description="Archive and index a downloaded construction-standard PDF.")
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--number")
    parser.add_argument("--title")
    parser.add_argument("--source-url")
    parser.add_argument("--retrieved-on", default=date.today().isoformat())
    parser.add_argument("--status", default="retrieved")
    parser.add_argument("--project", action="append", default=[])
    parser.add_argument("--notes")
    parser.add_argument("--library-root", type=Path)
    parser.add_argument("--refresh-metadata", action="store_true")
    args = parser.parse_args()

    source = args.pdf.expanduser().resolve()
    if not source.exists() or source.suffix.lower() != ".pdf":
        raise SystemExit(f"Valid PDF required: {source}")

    try:
        page_count = len(PdfReader(str(source)).pages)
    except Exception as exc:
        raise SystemExit(f"Unreadable PDF: {source}: {exc}") from exc
    if page_count < 1:
        raise SystemExit(f"PDF has no pages: {source}")

    skill_root = Path(__file__).resolve().parents[1]
    library_root = (args.library_root or (skill_root / "standards")).expanduser().resolve()
    library = library_root / "library"
    index_path = library_root / "index.json"
    library.mkdir(parents=True, exist_ok=True)

    digest = sha256(source)
    data = load_index(index_path)
    existing = next((item for item in data["standards"] if item.get("sha256") == digest), None)

    number = args.number or infer_number(source.stem)
    title = args.title or re.sub(r"^[^_]+_?", "", source.stem).strip() or source.stem
    source_record = {
        "url": args.source_url or "local-file",
        "retrieved_on": args.retrieved_on,
    }

    if existing:
        if args.refresh_metadata:
            existing["number"] = number
            existing["status"] = args.status
            if args.title:
                existing["title"] = title
        append_unique(existing.setdefault("projects", []), None)
        for project in args.project:
            append_unique(existing.setdefault("projects", []), project)
        if source_record not in existing.setdefault("sources", []):
            existing["sources"].append(source_record)
        if args.notes:
            append_unique(existing.setdefault("notes", []), args.notes)
        write_index(index_path, data)
        print(json.dumps({"action": "deduplicated", "path": str(library / existing["filename"]), "sha256": digest, "pages": page_count}, ensure_ascii=False))
        return

    filename = safe_filename(source.name)
    destination = library / filename
    if destination.exists() and sha256(destination) != digest:
        destination = library / f"{destination.stem}__{digest[:8]}{destination.suffix}"
    if source != destination:
        shutil.copy2(source, destination)

    record = {
        "number": number,
        "title": title,
        "filename": destination.name,
        "sha256": digest,
        "size_bytes": destination.stat().st_size,
        "pages": page_count,
        "status": args.status,
        "sources": [source_record],
        "projects": sorted(set(args.project)),
        "notes": [args.notes] if args.notes else [],
    }
    data["standards"].append(record)
    write_index(index_path, data)
    print(json.dumps({"action": "archived", "path": str(destination), "sha256": digest, "pages": page_count}, ensure_ascii=False))


if __name__ == "__main__":
    main()
