from __future__ import annotations

from copy import deepcopy
from hashlib import sha256
import json
from pathlib import Path
import re
import shutil
import zipfile

from docx import Document
from docx.enum.table import (
    WD_CELL_VERTICAL_ALIGNMENT,
    WD_ROW_HEIGHT_RULE,
    WD_TABLE_ALIGNMENT,
)
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.shared import Cm, Pt, RGBColor


ROOT = Path(__file__).resolve().parent
SOURCE = Path(
    r"C:\Users\Administrator\xwechat_files\wxid_bspmv0rjp1rf22_3149\msg\file\2026-08\条材厂型材除鳞系统改造施工方案(3333)(1).docx"
)
OUTPUT = ROOT / "output" / "条材厂型材除鳞系统改造施工方案_规范原文修订版.docx"
ASSETS = ROOT / "assets"

BODY_FONT = "仿宋_GB2312"
COVER_FONT = "方正仿宋_GB2312"
COMPANY = "武汉兴达建筑安装工程有限责任公司"
REVISION_DATE = "2026年8月5日"


def file_hash(path: Path) -> str:
    digest = sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest().upper()


def set_run_style(run, size=12, bold=False, font=BODY_FONT):
    run.font.name = font
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), font)
    rfonts.set(qn("w:hAnsi"), font)
    rfonts.set(qn("w:eastAsia"), font)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)


def clear_paragraph_text_keep_first_run(paragraph: Paragraph, text: str):
    if not paragraph.runs:
        run = paragraph.add_run(text)
        return run
    paragraph.runs[0].text = text
    for run in paragraph.runs[1:]:
        run.text = ""
    return paragraph.runs[0]


def replace_once_across_runs(paragraph: Paragraph, old: str, new: str) -> bool:
    text = paragraph.text
    start = text.find(old)
    if start < 0:
        return False
    end = start + len(old)
    runs = paragraph.runs
    cursor = 0
    start_run = start_offset = end_run = end_offset = None
    for index, run in enumerate(runs):
        next_cursor = cursor + len(run.text)
        if start_run is None and start < next_cursor:
            start_run = index
            start_offset = start - cursor
        if end_run is None and end <= next_cursor:
            end_run = index
            end_offset = end - cursor
            break
        cursor = next_cursor
    if start_run is None or end_run is None:
        return False
    if start_run == end_run:
        run = runs[start_run]
        run.text = run.text[:start_offset] + new + run.text[end_offset:]
        return True
    prefix = runs[start_run].text[:start_offset]
    suffix = runs[end_run].text[end_offset:]
    runs[start_run].text = prefix + new
    for index in range(start_run + 1, end_run):
        runs[index].text = ""
    runs[end_run].text = suffix
    return True


def replace_all(paragraph: Paragraph, old: str, new: str):
    if old in new:
        if old in paragraph.text and not replace_once_across_runs(paragraph, old, new):
            raise RuntimeError(f"Run-level replacement failed: {old}")
        return
    while old in paragraph.text:
        if not replace_once_across_runs(paragraph, old, new):
            raise RuntimeError(f"Run-level replacement failed: {old}")


def find_exact(document: Document, text: str, occurrence=-1) -> Paragraph:
    matches = [p for p in document.paragraphs if p.text.strip() == text]
    if not matches:
        raise ValueError(f"Paragraph not found: {text}")
    return matches[occurrence]


def find_contains(document: Document, text: str, occurrence=0) -> Paragraph:
    matches = [p for p in document.paragraphs if text in p.text]
    if not matches:
        raise ValueError(f"Paragraph containing text not found: {text}")
    return matches[occurrence]


def new_paragraph_after(anchor: Paragraph) -> Paragraph:
    element = OxmlElement("w:p")
    anchor._p.addnext(element)
    return Paragraph(element, anchor._parent)


def new_paragraph_before(anchor: Paragraph) -> Paragraph:
    element = OxmlElement("w:p")
    anchor._p.addprevious(element)
    return Paragraph(element, anchor._parent)


def format_text_paragraph(
    paragraph: Paragraph,
    text: str,
    *,
    bold=False,
    indent=True,
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    size=12,
    keep_with_next=False,
    page_break_before=False,
):
    paragraph.style = "Normal"
    paragraph.alignment = align
    fmt = paragraph.paragraph_format
    fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.keep_with_next = keep_with_next
    fmt.page_break_before = page_break_before
    fmt.left_indent = Pt(0)
    fmt.right_indent = Pt(0)
    fmt.first_line_indent = Pt(24) if indent else Pt(0)
    run = paragraph.add_run(text)
    set_run_style(run, size=size, bold=bold)
    return paragraph


def insert_body_after(anchor: Paragraph, text: str, **kwargs) -> Paragraph:
    paragraph = new_paragraph_after(anchor)
    return format_text_paragraph(paragraph, text, **kwargs)


def insert_body_before(anchor: Paragraph, text: str, **kwargs) -> Paragraph:
    paragraph = new_paragraph_before(anchor)
    return format_text_paragraph(paragraph, text, **kwargs)


def insert_figure_after(
    anchor: Paragraph,
    image_name: str,
    caption: str,
    *,
    width_cm: float,
    page_break_before=False,
) -> Paragraph:
    image_para = new_paragraph_after(anchor)
    image_para.style = "Normal"
    image_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    image_para.paragraph_format.space_before = Pt(3)
    image_para.paragraph_format.space_after = Pt(2)
    image_para.paragraph_format.keep_with_next = True
    image_para.paragraph_format.page_break_before = page_break_before
    run = image_para.add_run()
    inline = run.add_picture(str(ASSETS / image_name), width=Cm(width_cm))
    inline._inline.docPr.set("descr", caption)

    caption_para = new_paragraph_after(image_para)
    caption_para.style = "Normal"
    caption_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption_para.paragraph_format.space_before = Pt(0)
    caption_para.paragraph_format.space_after = Pt(4)
    caption_para.paragraph_format.keep_with_next = False
    caption_run = caption_para.add_run(caption)
    set_run_style(caption_run, size=9, bold=False)
    return caption_para


def set_cell_margins(cell, top=70, start=80, bottom=70, end=80):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table, size=6):
    tbl_pr = table._tbl.tblPr
    old = tbl_pr.find(qn("w:tblBorders"))
    if old is not None:
        tbl_pr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), str(size))
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "000000")
        borders.append(border)
    tbl_pr.append(borders)


def set_cell_shading(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    old = tc_pr.find(qn("w:shd"))
    if old is not None:
        tc_pr.remove(old)
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), fill)
    tc_pr.append(shading)


def set_cell_width(cell, width_cm: float):
    cell.width = Cm(width_cm)
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_width = tc_pr.find(qn("w:tcW"))
    if tc_width is None:
        tc_width = OxmlElement("w:tcW")
        tc_pr.append(tc_width)
    tc_width.set(qn("w:w"), str(round(width_cm / 2.54 * 1440)))
    tc_width.set(qn("w:type"), "dxa")


def set_cell_no_wrap(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    if tc_pr.find(qn("w:noWrap")) is None:
        tc_pr.append(OxmlElement("w:noWrap"))


def set_fixed_table_layout(table, width_cm: float):
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_pr = table._tbl.tblPr
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")
    width = tbl_pr.find(qn("w:tblW"))
    if width is None:
        width = OxmlElement("w:tblW")
        tbl_pr.append(width)
    width.set(qn("w:w"), str(round(width_cm / 2.54 * 1440)))
    width.set(qn("w:type"), "dxa")
    indent = tbl_pr.find(qn("w:tblInd"))
    if indent is None:
        indent = OxmlElement("w:tblInd")
        tbl_pr.append(indent)
    indent.set(qn("w:w"), "0")
    indent.set(qn("w:type"), "dxa")


def set_table_grid_widths(table, widths_cm):
    grid = table._tbl.tblGrid
    grid_columns = grid.findall(qn("w:gridCol"))
    if len(grid_columns) != len(widths_cm):
        raise RuntimeError("Table grid does not match the requested column widths")
    for grid_column, width_cm in zip(grid_columns, widths_cm):
        grid_column.set(qn("w:w"), str(round(width_cm / 2.54 * 1440)))


def repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:tblHeader")) is None:
        node = OxmlElement("w:tblHeader")
        node.set(qn("w:val"), "true")
        tr_pr.append(node)


def prevent_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    if tr_pr.find(qn("w:cantSplit")) is None:
        tr_pr.append(OxmlElement("w:cantSplit"))


def set_continuous_page_numbers(document: Document):
    footer = document.sections[0].footer
    paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)

    begin_run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin_run._r.append(begin)

    instruction_run = paragraph.add_run()
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    instruction_run._r.append(instruction)

    separator_run = paragraph.add_run()
    separator = OxmlElement("w:fldChar")
    separator.set(qn("w:fldCharType"), "separate")
    separator_run._r.append(separator)

    result_run = paragraph.add_run("1")
    set_run_style(result_run, size=9, bold=False)

    end_run = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run._r.append(end)

    for index, section in enumerate(document.sections):
        if index:
            section.footer.is_linked_to_previous = True
        sect_pr = section._sectPr
        pg_num_type = sect_pr.find(qn("w:pgNumType"))
        if pg_num_type is None:
            pg_num_type = OxmlElement("w:pgNumType")
            columns = sect_pr.find(qn("w:cols"))
            if columns is None:
                sect_pr.append(pg_num_type)
            else:
                sect_pr.insert(list(sect_pr).index(columns), pg_num_type)
        pg_num_type.set(qn("w:fmt"), "decimal")
        if qn("w:start") in pg_num_type.attrib:
            del pg_num_type.attrib[qn("w:start")]


def table_text(table):
    return [[cell.text for cell in row.cells] for row in table.rows]


OUTPUT.parent.mkdir(parents=True, exist_ok=True)
source_hash_before = file_hash(SOURCE)
if OUTPUT.exists():
    OUTPUT.chmod(0o666)
    OUTPUT.unlink()
shutil.copyfile(SOURCE, OUTPUT)
document = Document(OUTPUT)

personnel_snapshot = {
    "labour_plan": table_text(document.tables[4]),
    "personnel_list": table_text(document.tables[5]),
}

# Match the retained template's A4 geometry while preserving the source's three-section structure.
for index, section in enumerate(document.sections):
    if index < 2:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.left_margin = Cm(3.175)
        section.right_margin = Cm(3.175)
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
    else:
        section.page_width = Cm(29.7)
        section.page_height = Cm(21.0)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)
        section.top_margin = Cm(3.175)
        section.bottom_margin = Cm(3.175)
    section.header_distance = Cm(0)
    section.footer_distance = Cm(1.75)

# Cover and front matter.
company_para = find_contains(document, "单位名称（盖章）")
company_run = clear_paragraph_text_keep_first_run(company_para, COMPANY)
set_run_style(company_run, size=14, font=COVER_FONT)
company_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
company_para.paragraph_format.first_line_indent = Pt(0)

date_para = find_contains(document, "2026 年 7月 5日")
date_run = clear_paragraph_text_keep_first_run(date_para, REVISION_DATE)
set_run_style(date_run, size=14, font=COVER_FONT)
date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
date_para.paragraph_format.first_line_indent = Pt(0)

toc_updates = {
    "三、施工预设目标": "三、施工预设目标及组织管理体系",
    "九、施工进度计划": "十、施工进度计划",
}
for old, new in toc_updates.items():
    paragraph = find_exact(document, old, occurrence=0)
    run = clear_paragraph_text_keep_first_run(paragraph, new)
    set_run_style(run, size=14, font=BODY_FONT)

toc_safety = find_exact(document, "七、安全环保措施", occurrence=0)
toc_hazard = insert_body_after(
    toc_safety,
    "危险因素分析及对策措施",
    indent=True,
    align=WD_ALIGN_PARAGRAPH.LEFT,
    size=14,
)
toc_hazard.paragraph_format.first_line_indent = Pt(28)
toc_hazard.paragraph_format.line_spacing = 2.0

toc_resource = find_exact(document, "八、资源需求计划", occurrence=0)
toc_emergency = insert_body_after(
    toc_resource,
    "九、应急预案",
    indent=True,
    align=WD_ALIGN_PARAGRAPH.LEFT,
    size=14,
)
toc_emergency.paragraph_format.first_line_indent = Pt(28)
toc_emergency.paragraph_format.line_spacing = 2.0

for paragraph_index, paragraph in enumerate(document.paragraphs):
    if paragraph in (toc_hazard, toc_emergency):
        continue
    if paragraph.text.strip() in {
        "一、编制依据及工程概况",
        "二、施工组织体系",
        "三、施工预设目标及组织管理体系",
        "四、主要施工工序及方法",
        "五、技术质量标准",
        "六、质量保证措施",
        "七、安全环保措施",
        "八、资源需求计划",
        "十、施工进度计划",
    } and paragraph_index < 60:
        paragraph.paragraph_format.line_spacing = 2.0
        paragraph.paragraph_format.first_line_indent = Pt(28)
        for run in paragraph.runs:
            if run.text:
                set_run_style(run, size=14, font=BODY_FONT)

# Capture figure anchors before run-level standard corrections.
anchor_basis = find_contains(document, "12、现场管理")
anchor_temp_power = find_contains(document, "JGJ/T 46-2024")
anchor_pipe = find_contains(document, "《工业金属管道工程施工规范》GB 50235-2010建标库")
anchor_cable = find_contains(document, "《电气装置安装工程电缆线路施工及验收标准》GB 50168-2018")
anchor_steel = find_contains(document, "《钢结构工程施工质量验收标准》GB 50205-2020")

# Correct current standard references without flattening paragraph/run formatting.
standard_replacements = [
    ("GB/T20801-2020", "GB/T 20801.1—2025"),
    ("GB/T 20801.1~6-2025", "GB/T 20801.1—2025"),
    ("GB 50252-2010", "GB/T 50252—2018"),
    ("GB 50131-2007", "GB 50093—2013"),
    ("JGJ/T 46-2024", "JGJ/T 46—2024"),
    ("GB 50235-2010建标库", "GB 50235—2010"),
    ("GB 50254-2014建标库", "GB 50254—2014"),
    ("《特种设备焊接操作人员考核细则》TSG Z6002", "《特种设备焊接操作人员考核细则》TSG Z6002—2026"),
    ("NB/T 47013（", "NB/T 47013.2—2015（含2018年第1号修改单；"),
    ("NB/T 47013、", "NB/T 47013.2—2015（含2018年第1号修改单）、"),
    ("NB/T 47013.2，", "NB/T 47013.2—2015（含2018年第1号修改单），"),
]
for paragraph in document.paragraphs:
    for old, new in standard_replacements:
        if old in paragraph.text:
            replace_all(paragraph, old, new)

# Align core chapter names with the retained company template.
chapter_org = find_exact(document, "施工组织体系", occurrence=-1)
org_run = clear_paragraph_text_keep_first_run(chapter_org, "二、施工组织体系")
set_run_style(org_run, size=12, bold=True)

chapter_goal = find_exact(document, "三、施工预设目标", occurrence=-1)
goal_run = clear_paragraph_text_keep_first_run(chapter_goal, "三、施工预设目标及组织管理体系")
set_run_style(goal_run, size=12, bold=True)

hazard_heading = find_contains(document, "3、各分项重大危险源及专项防控交底")
hazard_run = clear_paragraph_text_keep_first_run(hazard_heading, "危险因素分析及对策措施")
set_run_style(hazard_run, size=12, bold=True)

resource_duplicate = find_exact(document, "八、资源需求计划", occurrence=-1)
resource_run = clear_paragraph_text_keep_first_run(resource_duplicate, "（五）资源计划明细表")
set_run_style(resource_run, size=12, bold=True)

find_contains(document, "1、工机具计划").runs[0].text = "1、主要施工机具计划"
find_exact(document, "台班计划").runs[0].text = "2、主要机械台班计划"

# Standardize existing body typography to the retained template without touching front matter.
body_start_matches = [
    index
    for index, paragraph in enumerate(document.paragraphs)
    if paragraph.text.strip() == "一、编制依据及工程概况"
]
body_start = body_start_matches[-1]
main_heading_re = re.compile(r"^[一二三四五六七八九十]+、")
sub_heading_re = re.compile(r"^（[一二三四五六七八九十]+）")
number_heading_re = re.compile(r"^\d+(?:\.\d+){0,3}[、.]\s*")

for paragraph in document.paragraphs[body_start:]:
    text = paragraph.text.strip()
    xml = paragraph._p.xml
    if not text or "w:object" in xml or "w:drawing" in xml or "w:pict" in xml:
        continue
    is_main = bool(main_heading_re.match(text)) or text == "危险因素分析及对策措施"
    is_sub = bool(sub_heading_re.match(text))
    is_short_numbered = bool(number_heading_re.match(text)) and len(text) <= 34
    fmt = paragraph.paragraph_format
    fmt.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    fmt.space_after = Pt(0)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT if (is_main or is_sub or is_short_numbered) else WD_ALIGN_PARAGRAPH.JUSTIFY
    if is_main or is_sub or is_short_numbered:
        fmt.first_line_indent = Pt(0)
        fmt.space_before = Pt(5 if is_main else 2)
        fmt.keep_with_next = True
    else:
        fmt.space_before = Pt(0)
        if paragraph.style.name != "List Paragraph":
            fmt.first_line_indent = Pt(24)
    for run in paragraph.runs:
        if run.text:
            set_run_style(run, size=12, bold=(is_main or is_sub or is_short_numbered))

# Insert standard-text evidence from the archived PDF originals.
anchor = insert_body_after(
    anchor_basis,
    "（三）规范原文引用与版本适用说明",
    bold=True,
    indent=False,
    align=WD_ALIGN_PARAGRAPH.LEFT,
    keep_with_next=True,
)
anchor = insert_body_after(
    anchor,
    "本次修订以2026年8月5日为版本核对基准日。与本方案直接相关的标准原文已统一归档至项目配套规范库，以下插图均从标准PDF正文页直接截取，保留标准编号、条款编号和原规范页码，不采用网页查询画面代替条文证据。原稿涉及工业管道和特种设备焊接人员的部分依据已经换版，应按开工时已经实施的现行版本执行。",
)
anchor = insert_body_after(
    anchor,
    "本项目新增并统一采用《压力管道规范 第1部分：工业管道》（GB/T 20801.1—2025）、《工业管道安全技术规程》（TSG 31—2025）、《特种设备焊接操作人员考核细则》（TSG Z6002—2026）以及《建筑与市政工程施工现场临时用电安全技术标准》（JGJ/T 46—2024）。承压焊缝射线检测采用《承压设备无损检测 第2部分：射线检测》（NB/T 47013.2—2015，含2018年第1号修改单）。",
)

anchor = insert_figure_after(
    anchor,
    "规范原文_GBT20801.1-2025_焊接接头检查比例.png",
    "图1  GB/T 20801.1—2025表42焊接接头检查方法和比例（规范原文第112页，PDF第120页）",
    width_cm=14.2,
    page_break_before=True,
)
anchor = insert_body_after(
    anchor,
    "适用结论：GB/T 20801.1—2025已于2026年5月1日实施，并整体替代GB/T 20801.1～6—2020。本项目GC1高压工业管道的材料验收、预制安装、焊接、检验、试验和安全防护统一按2025版及设计文件执行，原稿中“GB/T 20801—2020”和“GB/T 20801.1～6—2025”的写法均已纠正。",
)
anchor = insert_figure_after(
    anchor,
    "规范原文_TSG31-2025_无损检测与耐压试验.png",
    "图2  TSG 31—2025第4.4.2.5～4.4.2.7条及第4.5.1条（规范原文第31页，PDF第35页）",
    width_cm=14.2,
    page_break_before=True,
)
anchor = insert_body_after(
    anchor,
    "适用结论：TSG 31—2025自2026年1月1日起施行，已整合修订原TSG D0001—2009等工业管道安全技术规范。施工期间应按该规程落实安装告知、焊接和无损检测质量控制、耐压试验隔离、监督检验资料和工程移交要求。",
)
anchor = insert_figure_after(
    anchor,
    "规范原文_TSGZ6002-2026_持证范围与有效期.png",
    "图3  TSG Z6002—2026第4.1～4.2条焊工持证范围与有效期（规范原文第7页，PDF第10页）",
    width_cm=14.2,
    page_break_before=True,
)
insert_body_after(
    anchor,
    "适用结论：TSG Z6002—2026自2026年8月1日起施行。项目使用的GC1压力管道焊工证书须处于有效期内，焊接方法、母材类别、试件位置和覆盖范围须与实际施焊项目匹配；焊工、焊缝编号、焊接工艺参数和无损检测结果应形成可追溯记录。",
)

# Insert compact, process-specific standards evidence beside the relevant clauses.
fig_anchor = insert_figure_after(
    anchor_temp_power,
    "规范原文_JGJT46-2024_三级配电与TNS.png",
    "图4  JGJ/T 46—2024第3.1.1～3.2.1条三级配电与TN-S系统（规范原文第7页，PDF第17页）",
    width_cm=14.2,
    page_break_before=True,
)
insert_body_after(
    fig_anchor,
    "本项目临时用电采用TN-S系统和三级配电、两级剩余电流保护，配电箱执行“一机、一闸、一漏、一箱”，电焊机、手持式电动工具和移动照明按作业环境设置相应保护；日检、周检和停复电记录纳入安全资料。",
)

fig_anchor = insert_figure_after(
    anchor_pipe,
    "规范原文_GB50235-2010_液压试验要求.png",
    "图5  GB 50235—2010第8.6.3～8.6.4条压力试验条件与液压试验要求（规范原文第42页，PDF第51页）",
    width_cm=14.2,
    page_break_before=True,
)
insert_body_after(
    fig_anchor,
    "GB 50235—2010用于管道元件和材料检验、坡口加工、组对安装、热处理、检查检验、压力试验、吹扫清洗和工程交接。本项目高压水管道在该规范基础上，同时执行GB/T 20801.1—2025和TSG 31—2025的更高要求。",
)

fig_anchor = insert_figure_after(
    anchor_cable,
    "规范原文_GB50168-2018_桥架转弯与接地.png",
    "图6  GB 50168—2018第5.2.4～5.2.10条桥架安装、转弯半径与接地（规范原文第10页，PDF第19页）",
    width_cm=14.2,
    page_break_before=True,
)
insert_body_after(
    fig_anchor,
    "电缆和光缆施工应控制运输保管、导管及支架安装、分层敷设、弯曲半径、牵引力、终端制作、标识和防火阻燃封堵；隐蔽前组织检查并留存绝缘测试、接线校核和电缆防火封堵验收记录。",
)

fig_anchor = insert_figure_after(
    anchor_steel,
    "规范原文_GB50205-2020_涂层缺陷修补.png",
    "图7  GB 50205—2020第13.2.6～13.3.1条涂层质量与缺陷修补（规范原文第98页，PDF第113页）",
    width_cm=14.2,
    page_break_before=True,
)
insert_body_after(
    fig_anchor,
    "设备支架、管道支吊架、检修平台和临时钢结构应按GB 50205—2020控制钢材与焊材复验、构件加工、焊缝外观、安装偏差、连接部位涂装和防腐层修补。涉及承压管道的焊接接头不得以普通钢结构焊缝验收代替压力管道专项检验。",
)

# Add the emergency chapter required by the retained template before the progress chapter.
schedule_heading = find_exact(document, "九、施工进度计划", occurrence=-1)
emergency_paragraphs = [
    ("九、应急预案", True, False),
    ("（一）应急组织及职责", True, False),
    ("项目经理为现场应急总指挥，技术负责人负责判断设备、管道和结构状态并提出处置边界，专职安全员负责警戒、疏散、清点和信息上报，各专业班组长负责切断本专业危险源并组织初期处置。发生事故时，现场人员必须服从统一指挥，不得擅自进入警戒区域。", False, True),
    ("（二）信息报告与响应程序", True, False),
    ("发现异常后立即停止作业，按“现场呼叫—切断能源—设置警戒—人员撤离—清点人数—逐级报告—专业处置”的顺序响应。报告内容包括时间、地点、作业内容、人员伤害、设备状态、介质压力、火情及已采取措施；未经确认不得擅自恢复施工。", False, True),
    ("（三）高压水系统泄漏、爆管和试压异常", True, False),
    ("立即停止泵组和试压设备，严禁人员正对法兰、阀门、焊口及临时堵板；确认泄压至零后方可接近。设置封闭警戒区，检查支撑、紧固和焊缝状态，查明原因并重新办理检验手续。带压紧固、敲击、补焊和拆卸一律禁止。", False, True),
    ("（四）起重吊装及机械伤害", True, False),
    ("发生吊物失稳、索具异常或设备倾斜时立即停止动作，保持警戒并疏散吊装半径内人员；不得在悬吊状态下强行调整。发生挤压或机械伤害时先停机、断能、可靠支撑，再实施救援，防止二次坠落或回转伤害。", False, True),
    ("（五）火灾、触电和动火异常", True, False),
    ("电气火灾应先切断电源，使用干粉灭火器处置，严禁带电用水灭火；触电事故应以绝缘工具使伤者脱离电源并立即开展心肺复苏。动火点周边出现异常烟气、可燃物复燃或监护条件失效时，立即停止动火并清理火源，确认复查合格后方可重新作业。", False, True),
    ("（六）射线探伤异常", True, False),
    ("发现无关人员误入辐射警戒区、警戒失效或设备异常时，探伤人员立即停止曝光并封锁现场，核对个人剂量和人员暴露情况，按规定上报并接受专业评估。未经专业确认，不得自行拆检射线装置或隐瞒异常记录。", False, True),
    ("（七）高处坠落、有限空间和职业伤害", True, False),
    ("高处坠落人员不得随意搬动，先固定颈椎和伤肢并联系医疗救援；有限空间出现缺氧、中毒迹象时，外部监护人立即启动强制通风和救援报告，救援人员必须佩戴相应呼吸防护，严禁盲目施救。化学品入眼或接触皮肤时持续清水冲洗并尽快就医。", False, True),
    ("（八）应急物资、演练与恢复", True, False),
    ("现场配置急救箱、干粉灭火器、防火毯、绝缘工具、救援绳、警戒带、应急照明和必要的检测仪器。开工前组织一次综合应急演练，高压试验、吊装和射线探伤前进行专项桌面推演。事故处置结束后，经建设单位、生产单位和项目技术安全人员联合确认，完成原因分析、隐患整改、复查销项和复工审批后方可恢复施工。", False, True),
]
for text, bold, indent in emergency_paragraphs:
    insert_body_before(
        schedule_heading,
        text,
        bold=bold,
        indent=indent,
        align=WD_ALIGN_PARAGRAPH.LEFT if bold else WD_ALIGN_PARAGRAPH.JUSTIFY,
        keep_with_next=bold,
    )
find_exact(document, "九、应急预案").paragraph_format.page_break_before = True

schedule_run = clear_paragraph_text_keep_first_run(schedule_heading, "十、施工进度计划")
set_run_style(schedule_run, size=12, bold=True)

duration_para = find_contains(document, "计划工期：34个绝对工作日")
duration_run = clear_paragraph_text_keep_first_run(
    duration_para,
    "计划总工期37日历天，其中实体施工及系统调试安排34天，预留3天用于建设单位生产窗口衔接、联合验收和不可预见事项；具体开竣工日期以建设单位批准的开工令为准。",
)
set_run_style(duration_run, size=12)
duration_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
duration_para.paragraph_format.first_line_indent = Pt(24)

# Clean and stabilize all data tables without changing any personnel content.
for table_index, table in enumerate(document.tables):
    if table_index > 0:
        set_table_borders(table)
        repeat_header(table.rows[0])
    font_size = 12 if table_index == 0 else 10.5 if table_index < 5 else 8.5
    for row_index, row in enumerate(table.rows):
        prevent_split(row)
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.line_spacing = 1.0
                for run in paragraph.runs:
                    if run.text:
                        set_run_style(run, size=font_size, bold=(row_index == 0 and table_index > 0))

# Refine the final personnel roster while preserving every source value.
personnel_table = document.tables[5]
personnel_widths = (0.90, 1.05, 0.90, 0.90, 1.35, 3.35, 3.65, 2.55)
set_fixed_table_layout(personnel_table, sum(personnel_widths))
set_table_grid_widths(personnel_table, personnel_widths)
set_table_borders(personnel_table, size=7)
repeat_header(personnel_table.rows[0])
for row_index, row in enumerate(personnel_table.rows):
    row.height = Cm(0.72 if row_index == 0 else 0.62)
    row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    prevent_split(row)
    for column_index, (cell, width_cm) in enumerate(zip(row.cells, personnel_widths)):
        set_cell_width(cell, width_cm)
        set_cell_margins(cell, top=58, start=55, bottom=58, end=55)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        if row_index == 0:
            set_cell_shading(cell, "D9E2F3")
        elif row_index % 2 == 0:
            set_cell_shading(cell, "F5F7FA")
        else:
            set_cell_shading(cell, "FFFFFF")
        if row_index == 0 or column_index in (5, 6, 7):
            set_cell_no_wrap(cell)
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.keep_with_next = False
            for run in paragraph.runs:
                if not run.text:
                    continue
                size = 7.5 if row_index == 0 or column_index in (5, 6, 7) else 8.5 if column_index == 1 else 8
                set_run_style(run, size=size, bold=(row_index == 0))

# Keep the personnel table, emergency plan and progress narrative in the portrait
# body section; reserve the landscape section for the embedded progress chart only.
body = document._element.body
last_ole = None
last_section_break_para = None
for paragraph in document.paragraphs:
    if "o:OLEObject" in paragraph._p.xml:
        last_ole = paragraph._p
    if "w:sectPr" in paragraph._p.xml:
        last_section_break_para = paragraph
if last_ole is not None and last_section_break_para is not None:
    old_p_pr = last_section_break_para._p.get_or_add_pPr()
    section_properties = old_p_pr.find(qn("w:sectPr"))
    if section_properties is not None:
        old_p_pr.remove(section_properties)
        preceding = last_ole.getprevious()
        while preceding is not None and preceding.tag == qn("w:p"):
            preceding_paragraph = Paragraph(preceding, document._body)
            if preceding_paragraph.text.strip() or "w:sectPr" in preceding.xml:
                break
            previous = preceding.getprevious()
            body.remove(preceding)
            preceding = previous
        if preceding is None or preceding.tag != qn("w:p"):
            raise RuntimeError("Unable to place the landscape section break before the chart")
        preceding.get_or_add_pPr().append(section_properties)

# Remove only the trailing blank paragraphs after the final embedded progress chart.
if last_ole is not None:
    sibling = last_ole.getnext()
    while sibling is not None and sibling.tag == qn("w:p"):
        next_sibling = sibling.getnext()
        if "w:sectPr" not in sibling.xml:
            body.remove(sibling)
        sibling = next_sibling

# Verify the exact personnel configuration and list before saving.
assert table_text(document.tables[4]) == personnel_snapshot["labour_plan"]
assert table_text(document.tables[5]) == personnel_snapshot["personnel_list"]

settings = document.settings._element
update_fields = settings.find(qn("w:updateFields"))
if update_fields is None:
    update_fields = OxmlElement("w:updateFields")
    settings.append(update_fields)
update_fields.set(qn("w:val"), "true")

document.core_properties.title = "条材厂型材除鳞系统改造施工方案"
document.core_properties.subject = "按条材厂施工方案模板修订，附规范原文条款截图"
document.core_properties.author = COMPANY
set_continuous_page_numbers(document)
document.save(OUTPUT)

with zipfile.ZipFile(OUTPUT) as archive:
    bad_member = archive.testzip()

final_document = Document(OUTPUT)
validation = {
    "source_sha256_before": source_hash_before,
    "source_sha256_after": file_hash(SOURCE),
    "source_unchanged": source_hash_before == file_hash(SOURCE),
    "output": str(OUTPUT),
    "paragraphs": len(final_document.paragraphs),
    "tables": len(final_document.tables),
    "inline_shapes": len(final_document.inline_shapes),
    "sections": len(final_document.sections),
    "personnel_unchanged": table_text(final_document.tables[4]) == personnel_snapshot["labour_plan"]
    and table_text(final_document.tables[5]) == personnel_snapshot["personnel_list"],
    "zip_bad_member": bad_member,
}
(ROOT / "qa" / "build_validation.json").write_text(
    json.dumps(validation, ensure_ascii=False, indent=2), encoding="utf-8"
)
print(json.dumps(validation, ensure_ascii=False, indent=2))
