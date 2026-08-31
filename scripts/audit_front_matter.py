#!/usr/bin/env python3
import argparse
import copy
import hashlib
import json
import re
import zipfile
from pathlib import Path

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W_NS}
VOLATILE_ATTRS = {
    f"{{{W14_NS}}}paraId",
    f"{{{W14_NS}}}textId",
}
IGNORED_ELEMENTS = {
    f"{{{W_NS}}}proofErr",
    f"{{{W_NS}}}bookmarkStart",
    f"{{{W_NS}}}bookmarkEnd",
}


def read_part(path, part):
    with zipfile.ZipFile(path) as archive:
        return archive.read(part)


def xml_root(path, part):
    return etree.fromstring(read_part(path, part))


def local_name(element):
    return etree.QName(element).localname


def visible_text(element):
    return "".join(element.xpath(".//w:t/text()", namespaces=NS)).strip()


def cell_margin_signature(element):
    if element is None:
        return None
    return {
        local_name(child): (
            child.get(f"{{{W_NS}}}w"),
            child.get(f"{{{W_NS}}}type"),
        )
        for child in element
    }


def remove_redundant_row_cell_margins(root):
    """Ignore Word's duplicate row-level cell margins when they equal the table defaults."""
    for row_properties in root.xpath(".//w:tblPrEx", namespaces=NS):
        row_margins = row_properties.find(f"{{{W_NS}}}tblCellMar")
        if row_margins is None:
            continue
        table = next(
            (ancestor for ancestor in row_properties.iterancestors() if local_name(ancestor) == "tbl"),
            None,
        )
        if table is None:
            continue
        table_margins = table.find(f"./{{{W_NS}}}tblPr/{{{W_NS}}}tblCellMar")
        if cell_margin_signature(row_margins) == cell_margin_signature(table_margins):
            row_properties.remove(row_margins)


def normalize_xml(element):
    clone = copy.deepcopy(element)
    remove_redundant_row_cell_margins(clone)
    for item in list(clone.iter()):
        for attribute in list(item.attrib):
            qname = etree.QName(attribute)
            if attribute in VOLATILE_ATTRS or qname.localname.startswith("rsid"):
                del item.attrib[attribute]
        if item.tag in IGNORED_ELEMENTS:
            parent = item.getparent()
            if parent is not None:
                parent.remove(item)
                continue
        if item.tag in {
            f"{{{W_NS}}}t",
            f"{{{W_NS}}}delText",
            f"{{{W_NS}}}instrText",
        }:
            item.text = ""
            xml_space = "{http://www.w3.org/XML/1998/namespace}space"
            item.attrib.pop(xml_space, None)
    return etree.tostring(clone, method="c14n", exclusive=True, with_comments=False)


def digest(element):
    return hashlib.sha256(normalize_xml(element)).hexdigest()


def body_prefix(path, marker, occurrence):
    root = xml_root(path, "word/document.xml")
    body = root.find(f"{{{W_NS}}}body")
    blocks = [child for child in body if local_name(child) in {"p", "tbl"}]
    hits = 0
    boundary = None
    for index, block in enumerate(blocks):
        if marker in visible_text(block):
            hits += 1
            if hits == occurrence:
                boundary = index
                break
    if boundary is None:
        raise ValueError(f"body marker not found {occurrence} times: {marker!r}; found {hits}")
    return blocks[:boundary], boundary


def front_style_ids(blocks):
    values = {"Normal"}
    for block in blocks:
        for node in block.xpath(".//w:pStyle | .//w:rStyle | .//w:tblStyle", namespaces=NS):
            value = node.get(f"{{{W_NS}}}val")
            if value:
                values.add(value)
    return values


def style_map(path, requested_ids):
    root = xml_root(path, "word/styles.xml")
    styles = {
        node.get(f"{{{W_NS}}}styleId"): node
        for node in root.xpath("./w:style", namespaces=NS)
        if node.get(f"{{{W_NS}}}styleId")
    }
    selected = set(requested_ids)
    changed = True
    while changed:
        changed = False
        for style_id in list(selected):
            node = styles.get(style_id)
            if node is None:
                continue
            for tag in ("basedOn", "link", "next"):
                child = node.find(f"{{{W_NS}}}{tag}")
                if child is not None:
                    target = child.get(f"{{{W_NS}}}val")
                    if target and target not in selected:
                        selected.add(target)
                        changed = True
    result = {}
    defaults = root.find(f"{{{W_NS}}}docDefaults")
    result["__docDefaults__"] = hashlib.sha256(normalize_xml(defaults)).hexdigest() if defaults is not None else None
    for style_id in sorted(selected):
        node = styles.get(style_id)
        result[style_id] = hashlib.sha256(normalize_xml(node)).hexdigest() if node is not None else None
    return result


def first_section_map(path):
    result = {}
    with zipfile.ZipFile(path) as archive:
        document = etree.fromstring(archive.read("word/document.xml"))
        sections = document.xpath(".//w:sectPr", namespaces=NS)
        if not sections:
            return {"__section_properties__": None}
        section = sections[0]
        section_copy = copy.deepcopy(section)
        for reference in section_copy.xpath("./w:headerReference | ./w:footerReference", namespaces=NS):
            section_copy.remove(reference)
        result["__section_properties__"] = hashlib.sha256(normalize_xml(section_copy)).hexdigest()

        relationships = etree.fromstring(archive.read("word/_rels/document.xml.rels"))
        targets = {
            node.get("Id"): node.get("Target")
            for node in relationships.findall(f"{{{PKG_REL_NS}}}Relationship")
        }
        for reference in section.xpath("./w:headerReference | ./w:footerReference", namespaces=NS):
            relationship_id = reference.get(f"{{{R_NS}}}id")
            target = targets.get(relationship_id)
            kind = local_name(reference).replace("Reference", "")
            variant = reference.get(f"{{{W_NS}}}type", "default")
            key = f"{kind}:{variant}"
            if not target:
                result[key] = None
                continue
            part = "word/" + target.lstrip("/")
            root = etree.fromstring(archive.read(part))
            result[key] = hashlib.sha256(normalize_xml(root)).hexdigest()
    return result


def compare_maps(label, template_map, output_map, failures):
    differences = []
    for key in sorted(set(template_map) | set(output_map)):
        if template_map.get(key) != output_map.get(key):
            differences.append(
                {
                    "item": key,
                    "template": template_map.get(key),
                    "output": output_map.get(key),
                }
            )
    if differences:
        failures.append(f"{label} formatting differs in {len(differences)} item(s)")
    return differences


DEFAULT_ALLOWED_TEXT_PATTERNS = [
    r"^工程名称\s*[:：]",
    r"^工程编号\s*[:：]",
    r"^(?:【编制日期】|[〇零一二三四五六七八九十0-9]{4}年.*)$",
]


def text_change_allowed(template_block, output_block, in_contents, extra_patterns):
    if local_name(template_block) != "p" or local_name(output_block) != "p":
        return False
    if in_contents:
        return True
    before = visible_text(template_block)
    after = visible_text(output_block)
    before_signatory = re.match(r"^(编制|审核|审批|批准)\s*[:：]", before)
    after_signatory = re.match(r"^(编制|审核|审批|批准)\s*[:：]", after)
    if before_signatory or after_signatory:
        return bool(
            before_signatory
            and after_signatory
            and before_signatory.group(1) == after_signatory.group(1)
        )
    patterns = DEFAULT_ALLOWED_TEXT_PATTERNS + list(extra_patterns)
    return any(re.search(pattern, before) or re.search(pattern, after) for pattern in patterns)


def main():
    parser = argparse.ArgumentParser(description="Compare frozen front-matter formatting against the selected DOCX template.")
    parser.add_argument("template", type=Path)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--body-start", default="一、编制依据及工程概况")
    parser.add_argument("--occurrence", type=int, default=2)
    parser.add_argument("--allow-text-pattern", action="append", default=[])
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    template = args.template.expanduser().resolve()
    output = args.docx.expanduser().resolve()
    failures = []

    try:
        template_blocks, template_boundary = body_prefix(template, args.body_start, args.occurrence)
        output_blocks, output_boundary = body_prefix(output, args.body_start, args.occurrence)
    except Exception as error:
        result = {"template": str(template), "docx": str(output), "failures": [str(error)]}
        rendered = json.dumps(result, ensure_ascii=False, indent=2)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(rendered, encoding="utf-8")
        print(rendered)
        raise SystemExit(1)

    block_differences = []
    if len(template_blocks) != len(output_blocks):
        failures.append(f"frozen block count differs: {len(output_blocks)} != {len(template_blocks)}")
    for index in range(max(len(template_blocks), len(output_blocks))):
        template_block = template_blocks[index] if index < len(template_blocks) else None
        output_block = output_blocks[index] if index < len(output_blocks) else None
        if template_block is None or output_block is None:
            block_differences.append(
                {
                    "block": index + 1,
                    "template_type": local_name(template_block) if template_block is not None else None,
                    "output_type": local_name(output_block) if output_block is not None else None,
                    "template_text": visible_text(template_block)[:160] if template_block is not None else None,
                    "output_text": visible_text(output_block)[:160] if output_block is not None else None,
                }
            )
            continue
        if local_name(template_block) != local_name(output_block) or digest(template_block) != digest(output_block):
            block_differences.append(
                {
                    "block": index + 1,
                    "template_type": local_name(template_block),
                    "output_type": local_name(output_block),
                    "template_text": visible_text(template_block)[:160],
                    "output_text": visible_text(output_block)[:160],
                    "template_format_hash": digest(template_block),
                    "output_format_hash": digest(output_block),
                }
            )
    if block_differences:
        failures.append(f"frozen front-matter block formatting differs in {len(block_differences)} block(s)")

    style_ids = front_style_ids(template_blocks) | front_style_ids(output_blocks)
    style_differences = compare_maps(
        "front-matter styles",
        style_map(template, style_ids),
        style_map(output, style_ids),
        failures,
    )
    first_section_differences = compare_maps(
        "first-section settings and headers/footers",
        first_section_map(template),
        first_section_map(output),
        failures,
    )

    approved_text_changes = []
    unexpected_text_changes = []
    in_contents = False
    for index, (template_block, output_block) in enumerate(zip(template_blocks, output_blocks), start=1):
        before = visible_text(template_block)
        after = visible_text(output_block)
        if local_name(template_block) == "p" and re.sub(r"\s+", "", before) in {"目录", "目  录".replace(" ", "")}:
            in_contents = True
        if before != after:
            change = {"block": index, "template": before[:300], "output": after[:300]}
            if text_change_allowed(template_block, output_block, in_contents, args.allow_text_pattern):
                approved_text_changes.append(change)
            else:
                unexpected_text_changes.append(change)
    if unexpected_text_changes:
        failures.append(f"unexpected frozen front-matter text changes in {len(unexpected_text_changes)} block(s)")

    result = {
        "template": str(template),
        "docx": str(output),
        "body_start": args.body_start,
        "occurrence": args.occurrence,
        "template_boundary_block": template_boundary + 1,
        "output_boundary_block": output_boundary + 1,
        "frozen_blocks": len(template_blocks),
        "approved_text_changes_to_review": approved_text_changes,
        "unexpected_text_changes": unexpected_text_changes,
        "block_format_differences": block_differences,
        "style_differences": style_differences,
        "first_section_differences": first_section_differences,
        "failures": failures,
    }
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered, encoding="utf-8")
    print(rendered)
    raise SystemExit(1 if failures else 0)


if __name__ == "__main__":
    main()
