#!/usr/bin/env python3
"""Read-only audit of the user-adopted 2026-10-23 bidding format profile.

The source recommends Songti/12 pt/1.5 lines and says normally <=200 pages.
This audits the adopted production baseline, not a universal legal requirement.
Only an actual rendered PDF establishes page count; app.xml is never used.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
DEFAULT_PROFILE = Path(__file__).resolve().parents[1] / "references" / "latest-format-20261023.json"
PLACEHOLDERS = ("【填写：", "【填写:", "【项目名称】", "【工期】")


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def text_of(element) -> str:
    return "".join(node.text or "" for node in element.iter(W + "t"))


def normalized(text: str) -> str:
    return re.sub(r"[\s\u3000:：、，,。．.（）()\-—]+", "", text)


def bare_title(text: str) -> str:
    text = re.sub(r"^\s*(?:第[一二三四五六七八九十百\d]+章|附[表录][一二三四五六七八九十\d]+|[一二三四五六七八九十百\d]+[、.．])\s*[:：、.．]?", "", text)
    return normalized(text)


def title_matches(text: str, expected: str) -> bool:
    return normalized(text) == normalized(expected) or bare_title(text) == bare_title(expected)


class Styles:
    def __init__(self, root, theme=None):
        self.styles = {s.get(W + "styleId"): s for s in root.findall(W + "style")}
        self.defaults = root.find(W + "docDefaults")
        self.theme = theme

    def chain(self, style_id: str | None):
        result, seen = [], set()
        while style_id and style_id in self.styles and style_id not in seen:
            seen.add(style_id)
            style = self.styles[style_id]
            result.insert(0, style)
            base = style.find(W + "basedOn")
            style_id = base.get(W + "val") if base is not None else None
        return result

    def names(self, style_id):
        result = {style_id}
        for style in self.chain(style_id):
            result.add(style.get(W + "styleId"))
            name = style.find(W + "name")
            if name is not None:
                result.add(name.get(W + "val"))
        return result

    def theme_font(self, value):
        if self.theme is None or not value:
            return "unresolved-theme:" + str(value)
        group = "majorFont" if value.startswith("major") else "minorFont"
        scheme = self.theme.find(".//" + A + group)
        if scheme is not None:
            for font in scheme.findall(A + "font"):
                if font.get("script") == "Hans" and font.get("typeface"):
                    return font.get("typeface")
            ea = scheme.find(A + "ea")
            if ea is not None and ea.get("typeface"):
                return ea.get("typeface")
        return "unresolved-theme:" + str(value)

    def run_properties(self, paragraph, run):
        pstyle = paragraph.find(W + "pPr/" + W + "pStyle")
        style_id = pstyle.get(W + "val") if pstyle is not None else "Normal"
        layers = []
        if self.defaults is not None:
            layers.append(self.defaults.find(W + "rPrDefault/" + W + "rPr"))
        layers.extend(s.find(W + "rPr") for s in self.chain(style_id))
        rstyle = run.find(W + "rPr/" + W + "rStyle")
        if rstyle is not None:
            layers.extend(s.find(W + "rPr") for s in self.chain(rstyle.get(W + "val")))
        layers.append(run.find(W + "rPr"))
        result = {"font": None, "size_pt": None}
        for layer in layers:
            if layer is None:
                continue
            fonts = layer.find(W + "rFonts")
            if fonts is not None:
                theme = fonts.get(W + "eastAsiaTheme")
                explicit = fonts.get(W + "eastAsia")
                if theme:
                    result["font"] = self.theme_font(theme)
                elif explicit:
                    result["font"] = explicit
            size = layer.find(W + "sz")
            if size is not None:
                result["size_pt"] = float(size.get(W + "val")) / 2
        return result

    def spacing(self, paragraph):
        pstyle = paragraph.find(W + "pPr/" + W + "pStyle")
        style_id = pstyle.get(W + "val") if pstyle is not None else "Normal"
        layers = []
        if self.defaults is not None:
            layers.append(self.defaults.find(W + "pPrDefault/" + W + "pPr"))
        layers.extend(s.find(W + "pPr") for s in self.chain(style_id))
        layers.append(paragraph.find(W + "pPr"))
        result = {"line": "240", "lineRule": "auto"}
        for layer in layers:
            spacing = layer.find(W + "spacing") if layer is not None else None
            if spacing is not None:
                for key in result:
                    if spacing.get(W + key) is not None:
                        result[key] = spacing.get(W + key)
        return result


def paragraph_style(paragraph):
    style = paragraph.find(W + "pPr/" + W + "pStyle")
    return style.get(W + "val") if style is not None else "Normal"


def is_nonbody_role(paragraph, styles):
    names = styles.names(paragraph_style(paragraph))
    if any(re.match(r"(?:heading|biddingheading|biddingappendix|biddingcaption|biddingnote|biddingtitle|title|subtitle|caption|toc|biddingtoc)", str(n), re.I) for n in names):
        return True
    properties = [s.find(W + "pPr") for s in styles.chain(paragraph_style(paragraph))]
    properties.append(paragraph.find(W + "pPr"))
    outline = None
    for props in properties:
        level = props.find(W + "outlineLvl") if props is not None else None
        if level is not None:
            outline = level.get(W + "val")
    return outline is not None and outline.isdigit() and int(outline) < 9


def toc_indices(elements, styles):
    result, fields = set(), []
    for index, element in enumerate(elements):
        if element.tag != W + "p":
            continue
        in_toc = any(f["toc"] for f in fields)
        for node in element.iter():
            if node.tag == W + "fldChar":
                kind = node.get(W + "fldCharType")
                if kind == "begin":
                    fields.append({"instruction": "", "toc": False})
                elif kind == "end" and fields:
                    fields.pop()
            elif node.tag == W + "instrText" and fields:
                fields[-1]["instruction"] += node.text or ""
                fields[-1]["toc"] = bool(re.search(r"\bTOC\b", fields[-1]["instruction"], re.I))
            elif node.tag == W + "fldSimple" and re.search(r"\bTOC\b", node.get(W + "instr", ""), re.I):
                in_toc = True
            in_toc = in_toc or any(f["toc"] for f in fields)
        names = styles.names(paragraph_style(element))
        if in_toc or any(re.match(r"(?:toc|biddingtoc)", str(n), re.I) for n in names):
            result.add(index)
    return result


def header_columns(table):
    grid = table.find(W + "tblGrid")
    count = len(grid) if grid is not None else 0
    columns = [""] * count
    for row in table.findall(W + "tr")[:2]:
        logical = 0
        before = row.find(W + "trPr/" + W + "gridBefore")
        if before is not None:
            logical = int(before.get(W + "val", "0"))
        for cell in row.findall(W + "tc"):
            span_node = cell.find(W + "tcPr/" + W + "gridSpan")
            span = int(span_node.get(W + "val", "1")) if span_node is not None else 1
            for col in range(logical, min(logical + span, count)):
                columns[col] += text_of(cell)
            logical += span
    return count, columns


def audit(docx: Path, *, pdf: Path | None = None, profile: Path = DEFAULT_PROFILE,
          body_styles=None, allow_unrendered=False, allow_placeholders=False,
          page_limit_exception: str | None = None):
    """Return an audit report without modifying either input."""
    report = {"docx": str(docx), "profile": str(profile), "failures": [], "warnings": [],
              "policy": "User-adopted latest format; source font/spacing is recommended and page limit is in principle.",
              "release_ready": False, "clean_compliant": False}
    failures, warnings = report["failures"], report["warnings"]
    try:
        spec = json.loads(profile.read_text(encoding="utf-8-sig"))
        if len(spec["chapters"]) != 15 or len(spec["appendices"]) != 6:
            raise ValueError("profile must define exactly 15 chapters and 6 appendices")
        with zipfile.ZipFile(docx) as package:
            bad = package.testzip()
            if bad:
                raise ValueError(f"corrupt ZIP entry: {bad}")
            root = ET.fromstring(package.read("word/document.xml"))
            style_root = ET.fromstring(package.read("word/styles.xml"))
            theme = ET.fromstring(package.read("word/theme/theme1.xml")) if "word/theme/theme1.xml" in package.namelist() else None
            all_text = "\n".join(text_of(ET.fromstring(package.read(name))) for name in package.namelist()
                                 if name.startswith("word/") and name.endswith(".xml"))
        report["docx_sha256"] = digest(docx)
    except Exception as exc:
        failures.append(f"Input/profile could not be read: {exc}")
        report["status"] = "failed"
        return report

    styles = Styles(style_root, theme)
    elements = list(root.find(W + "body"))
    toc = toc_indices(elements, styles)
    paragraphs = [(i, p) for i, p in enumerate(elements) if p.tag == W + "p" and i not in toc]
    found = {"chapters": [], "appendices": []}
    title_locations = {}
    for kind in found:
        for entry in spec[kind]:
            title = entry["title"] if isinstance(entry, dict) else entry
            matches = [(i, p) for i, p in paragraphs if title_matches(text_of(p), title)]
            if not matches:
                failures.append(f"Missing body {kind[:-1]} heading: {title}")
            else:
                index = matches[0][0]
                found[kind].append({"title": title, "body_index": index})
                title_locations[(kind, title)] = index
    report["coverage"] = found
    chapter_locations = [item["body_index"] for item in found["chapters"]]
    if chapter_locations != sorted(chapter_locations):
        failures.append("Chapter headings are not in profile order")
    first_chapter = min(chapter_locations) if chapter_locations else None
    heading_indices = set(title_locations.values())
    for entry in spec["appendices"]:
        if not isinstance(entry, dict) or not entry.get("chapter"):
            continue
        anchor = title_locations.get(("appendices", entry["title"]))
        if anchor is None:
            continue
        preceding = [(title_locations[("chapters", ch["title"])], ch["number"])
                     for ch in spec["chapters"] if isinstance(ch, dict)
                     and ("chapters", ch["title"]) in title_locations
                     and title_locations[("chapters", ch["title"])] < anchor]
        actual_chapter = max(preceding)[1] if preceding else None
        if actual_chapter != entry["chapter"]:
            failures.append(f"Appendix placement mismatch: {entry['title']} is in chapter {actual_chapter}, expected {entry['chapter']}")

    selected = set(body_styles or ["BiddingBody"])
    report["body_scope"] = {"styles": sorted(selected), "starts_at_first_chapter": True}
    body_count, checked_runs, defects, uncovered = 0, 0, [], []
    for index, paragraph in paragraphs:
        if first_chapter is None or index < first_chapter or index in heading_indices:
            continue
        style_id = paragraph_style(paragraph)
        names = styles.names(style_id)
        text = text_of(paragraph).strip()
        if not text or is_nonbody_role(paragraph, styles):
            continue
        if not selected.intersection(names):
            if len(re.sub(r"\s", "", text)) >= 12:
                uncovered.append({"body_index": index, "style": style_id, "sample": text[:60]})
            continue
        body_count += 1
        spacing = styles.spacing(paragraph)
        if spacing["lineRule"] != "auto" or spacing["line"] != "360":
            defects.append({"body_index": index, "property": "line_spacing", "actual": spacing, "expected": "1.5 lines (auto 360)"})
        for run in paragraph.iter(W + "r"):
            text = text_of(run)
            if not text.strip():
                continue
            checked_runs += 1
            props = styles.run_properties(paragraph, run)
            if re.search(r"[\u3400-\u9fff]", text) and props["font"] not in {"宋体", "SimSun"}:
                defects.append({"body_index": index, "property": "east_asian_font", "actual": props["font"], "expected": "宋体", "sample": text[:30]})
            if props["size_pt"] != 12:
                defects.append({"body_index": index, "property": "font_size", "actual": props["size_pt"], "expected": 12, "sample": text[:30]})
    report.update(body_paragraphs_checked=body_count, text_runs_checked=checked_runs,
                  format_defects=defects, uncovered_substantive_paragraphs=uncovered)
    if uncovered:
        failures.append(f"Substantive body paragraphs outside audited styles: {len(uncovered)}; tag or explicitly include their body styles")
    if not body_count:
        failures.append("No body paragraphs in the selected style scope; format audit cannot pass")
    if defects:
        failures.append(f"Body formatting differs from adopted baseline: {len(defects)} defect(s)")

    table_checks = []
    for entry in spec["appendices"]:
        if not isinstance(entry, dict) or not (entry.get("columns") or entry.get("merged_header")):
            continue
        title, expected = entry["title"], entry.get("columns", [])
        anchor = title_locations.get(("appendices", title))
        if anchor is None:
            continue
        next_heading = min((i for i in heading_indices if i > anchor), default=len(elements))
        tables = [element for element in elements[anchor + 1:next_heading] if element.tag == W + "tbl"]
        if not tables:
            failures.append(f"Missing appendix table: {title}")
            continue
        count, headers = header_columns(tables[0])
        mismatches = [i + 1 for i, label in enumerate(expected)
                      if i >= len(headers) or normalized(str(label)) not in normalized(headers[i])]
        expected_count = entry.get("column_count", len(expected))
        passed = count == expected_count and not mismatches
        merged_header_passed = None
        if entry.get("merged_header"):
            first_row = tables[0].find(W + "tr")
            cells = first_row.findall(W + "tc") if first_row is not None else []
            span = cells[1].find(W + "tcPr/" + W + "gridSpan") if len(cells) > 1 else None
            merged_header_passed = (
                len(cells) == 2 and span is not None
                and int(span.get(W + "val", "1")) == expected_count - 1
                and normalized(entry.get("first_header", "")) in normalized(text_of(cells[0]))
                and normalized(entry["merged_header"]) in normalized(text_of(cells[1]))
            )
            passed = passed and merged_header_passed
        table_checks.append({"title": title, "columns": count, "expected_columns": expected_count,
                             "header_mismatch_columns": mismatches, "merged_header_passed": merged_header_passed,
                             "passed": passed})
        if not passed:
            failures.append(f"Appendix table header/schema mismatch: {title}")
    report["appendix_table_checks"] = table_checks

    markers = [marker for marker in PLACEHOLDERS if marker in all_text]
    report["placeholder_markers"] = markers
    if markers:
        (warnings if allow_placeholders else failures).append("Unfilled placeholders: " + ", ".join(markers))
    if allow_placeholders:
        warnings.append("Template QA mode: --allow-placeholders is not release validation")
    if pdf is None:
        (warnings if allow_unrendered else failures).append("Rendered PDF missing; actual page count is unverified")
        report["pages"] = None
    else:
        try:
            from pypdf import PdfReader
            count = len(PdfReader(str(pdf)).pages)
            if count < 1:
                raise ValueError("PDF has no pages")
            report.update(pdf=str(pdf), pdf_sha256=digest(pdf), pages=count)
            if count > 200:
                if page_limit_exception and page_limit_exception.strip():
                    report["page_limit_exception"] = page_limit_exception.strip()
                    warnings.append(f"Page-limit exception recorded: {count} pages; not clean <=200-page compliance")
                else:
                    failures.append(f"Page count {count} exceeds in-principle limit 200: needs review or explicit exception reason")
        except Exception as exc:
            failures.append(f"Rendered PDF could not be audited: {exc}")
            report["pages"] = None
    template_mode = allow_placeholders or allow_unrendered
    report["release_ready"] = not failures and not template_mode and pdf is not None
    report["clean_compliant"] = report["release_ready"] and "page_limit_exception" not in report
    if failures:
        report["status"] = "needs_review" if any("needs review" in f for f in failures) else "failed"
    elif template_mode:
        report["status"] = "template_qa_only"
    elif "page_limit_exception" in report:
        report["status"] = "passed_with_exception"
    else:
        report["status"] = "passed"
    return report


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--pdf", type=Path, help="Actual WPS/Word-rendered PDF, required for release")
    parser.add_argument("--profile", type=Path, default=DEFAULT_PROFILE)
    parser.add_argument("--body-style", action="append", help="Repeatable body style ID/name; default BiddingBody")
    parser.add_argument("--allow-unrendered", action="store_true", help="Template testing only; never release-ready")
    parser.add_argument("--allow-placeholders", action="store_true", help="Template testing only; never release-ready")
    parser.add_argument("--page-limit-exception", help="Explicitly approved reason; report remains exception-qualified")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.output and args.output.resolve() in {p.resolve() for p in (args.docx, args.pdf, args.profile) if p}:
        parser.error("--output must not overwrite any input")
    report = audit(args.docx, pdf=args.pdf, profile=args.profile, body_styles=args.body_style,
                   allow_unrendered=args.allow_unrendered, allow_placeholders=args.allow_placeholders,
                   page_limit_exception=args.page_limit_exception)
    content = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        args.output.write_text(content + "\n", encoding="utf-8")
    print(content)
    return 1 if report["failures"] else 0


if __name__ == "__main__":
    sys.exit(main())
