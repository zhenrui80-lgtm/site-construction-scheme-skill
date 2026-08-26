#!/usr/bin/env python3
"""Render every PDF page, create contact sheets, and flag suspicious layout pages."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageDraw, ImageFont, ImageStat


def ratio_below(image, threshold):
    histogram = image.histogram()
    return sum(histogram[:threshold]) / max(1, sum(histogram))


def page_metrics(image, page_number):
    gray = image.convert("L")
    stat = ImageStat.Stat(gray)
    border = 3
    strips = [
        gray.crop((0, 0, gray.width, border)),
        gray.crop((0, gray.height - border, gray.width, gray.height)),
        gray.crop((0, 0, border, gray.height)),
        gray.crop((gray.width - border, 0, gray.width, gray.height)),
    ]
    edge_pixels = sum(strip.width * strip.height for strip in strips)
    edge_ink = sum(sum(strip.histogram()[:245]) for strip in strips) / max(1, edge_pixels)
    return {
        "page": page_number,
        "width": image.width,
        "height": image.height,
        "orientation": "landscape" if image.width > image.height else "portrait",
        "mean_gray": round(stat.mean[0], 2),
        "stddev_gray": round(stat.stddev[0], 2),
        "ink_ratio": round(ratio_below(gray, 245), 6),
        "dark_ink_ratio": round(ratio_below(gray, 225), 6),
        "edge_ink_ratio": round(edge_ink, 6),
    }


def contact_sheets(page_paths, output_dir, columns=4, rows=4):
    contacts = []
    font = ImageFont.load_default()
    thumb_w, thumb_h, label_h = 245, 345, 22
    capacity = columns * rows
    for sheet_number, offset in enumerate(range(0, len(page_paths), capacity), start=1):
        canvas = Image.new("RGB", (columns * thumb_w, rows * (thumb_h + label_h)), "#d9dde2")
        draw = ImageDraw.Draw(canvas)
        for slot, path in enumerate(page_paths[offset: offset + capacity]):
            with Image.open(path).convert("RGB") as page:
                page.thumbnail((thumb_w - 10, thumb_h - 10), Image.Resampling.LANCZOS)
                x = (slot % columns) * thumb_w + (thumb_w - page.width) // 2
                y = (slot // columns) * (thumb_h + label_h) + 5
                canvas.paste(page, (x, y))
                draw.text((x, y + thumb_h), f"Page {offset + slot + 1}", fill="black", font=font)
        output = output_dir / f"contact-{sheet_number:02d}.jpg"
        canvas.save(output, quality=90)
        contacts.append(str(output))
    return contacts


def audit_pdf(
    pdf_path,
    output_dir,
    scale=1.5,
    allowed_sparse_pages=None,
    max_unexpected_sparse_pages=0,
    sparse_ink_ratio=0.025,
    blank_ink_ratio=0.001,
    max_edge_ink_ratio=0.02,
    expect_landscape_last_page=False,
):
    allowed_sparse_pages = set(allowed_sparse_pages or [])
    output_dir.mkdir(parents=True, exist_ok=True)
    pages_dir = output_dir / "pages"
    contacts_dir = output_dir / "contact_sheets"
    pages_dir.mkdir(parents=True, exist_ok=True)
    contacts_dir.mkdir(parents=True, exist_ok=True)
    document = pdfium.PdfDocument(str(pdf_path))
    metrics = []
    page_paths = []
    try:
        for index in range(len(document)):
            page = document[index]
            try:
                bitmap = page.render(scale=scale)
                try:
                    image = bitmap.to_pil().convert("RGB")
                finally:
                    bitmap.close()
                path = pages_dir / f"page-{index + 1:03d}.png"
                image.save(path, optimize=True)
                page_paths.append(path)
                metrics.append(page_metrics(image, index + 1))
            finally:
                page.close()
    finally:
        document.close()
    contacts = contact_sheets(page_paths, contacts_dir)
    blank_pages = [item["page"] for item in metrics if item["ink_ratio"] < blank_ink_ratio]
    sparse_pages = [item["page"] for item in metrics if item["ink_ratio"] < sparse_ink_ratio]
    unexpected_sparse = [page for page in sparse_pages if page not in allowed_sparse_pages]
    clipped_pages = [item["page"] for item in metrics if item["edge_ink_ratio"] > max_edge_ink_ratio]
    failures = []
    if blank_pages:
        failures.append(f"blank pages detected: {blank_pages}")
    if len(unexpected_sparse) > max_unexpected_sparse_pages:
        failures.append(
            f"unexpected sparse pages {unexpected_sparse} exceed allowed count {max_unexpected_sparse_pages}"
        )
    if clipped_pages:
        failures.append(f"possible edge clipping on pages: {clipped_pages}")
    if expect_landscape_last_page and (not metrics or metrics[-1]["orientation"] != "landscape"):
        failures.append("last page is not landscape as required for the schedule")
    return {
        "pdf": str(pdf_path),
        "page_count": len(metrics),
        "pages_dir": str(pages_dir),
        "contact_sheets": contacts,
        "thresholds": {
            "sparse_ink_ratio": sparse_ink_ratio,
            "blank_ink_ratio": blank_ink_ratio,
            "max_edge_ink_ratio": max_edge_ink_ratio,
            "max_unexpected_sparse_pages": max_unexpected_sparse_pages,
        },
        "allowed_sparse_pages": sorted(allowed_sparse_pages),
        "blank_pages": blank_pages,
        "sparse_pages": sparse_pages,
        "unexpected_sparse_pages": unexpected_sparse,
        "possible_clipped_pages": clipped_pages,
        "page_metrics": metrics,
        "failures": failures,
    }


def main():
    parser = argparse.ArgumentParser(description="Render and audit a final construction-scheme PDF.")
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--scale", type=float, default=1.5)
    parser.add_argument("--allowed-sparse-page", action="append", type=int, default=[])
    parser.add_argument("--max-unexpected-sparse-pages", type=int, default=0)
    parser.add_argument("--sparse-ink-ratio", type=float, default=0.025)
    parser.add_argument("--blank-ink-ratio", type=float, default=0.001)
    parser.add_argument("--max-edge-ink-ratio", type=float, default=0.02)
    parser.add_argument("--expect-landscape-last-page", action="store_true")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    result = audit_pdf(
        args.pdf.expanduser().resolve(),
        args.output_dir.expanduser().resolve(),
        args.scale,
        args.allowed_sparse_page,
        args.max_unexpected_sparse_pages,
        args.sparse_ink_ratio,
        args.blank_ink_ratio,
        args.max_edge_ink_ratio,
        args.expect_landscape_last_page,
    )
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.expanduser().resolve().write_text(rendered, encoding="utf-8")
    print(rendered)
    raise SystemExit(1 if result["failures"] else 0)


if __name__ == "__main__":
    main()
