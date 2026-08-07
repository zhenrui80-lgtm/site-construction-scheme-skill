#!/usr/bin/env python3
import argparse
import re
from pathlib import Path

import pypdfium2 as pdfium


def parse_pages(spec, total):
    pages = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        match = re.fullmatch(r"(\d+)(?:-(\d+))?", part)
        if not match:
            raise ValueError(f"Invalid page selection: {part}")
        start = int(match.group(1))
        end = int(match.group(2) or start)
        if start > end:
            start, end = end, start
        for page in range(start, end + 1):
            if page < 1 or page > total:
                raise ValueError(f"Page {page} outside PDF range 1-{total}")
            pages.add(page)
    return sorted(pages)


def safe_stem(value):
    return re.sub(r"[<>:\"/\\|?*]+", "_", value).strip(" .") or "standard"


def main():
    parser = argparse.ArgumentParser(description="Render exact pages from a downloaded standard PDF.")
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--pages", required=True, help="1-based pages, e.g. 3,7-9")
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--scale", type=float, default=2.0)
    parser.add_argument("--prefix")
    args = parser.parse_args()

    pdf_path = args.pdf.expanduser().resolve()
    if not pdf_path.exists():
        raise SystemExit(f"PDF not found: {pdf_path}")
    output_dir = args.output_dir.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    document = pdfium.PdfDocument(str(pdf_path))
    selected = parse_pages(args.pages, len(document))
    prefix = safe_stem(args.prefix or pdf_path.stem)
    outputs = []
    for page_number in selected:
        page = document[page_number - 1]
        image = page.render(scale=args.scale).to_pil().convert("RGB")
        output = output_dir / f"{prefix}_p{page_number:03d}.png"
        image.save(output, optimize=True)
        outputs.append(str(output))
    print(f"rendered={len(outputs)} total_pages={len(document)}")
    for output in outputs:
        print(output)


if __name__ == "__main__":
    main()
