---
name: site-construction-scheme
description: Create and revise full-depth professional Chinese metallurgical or industrial construction-scheme DOCX deliverables from BOQ spreadsheets, drawings, photos, existing schemes, and fixed schedules. Use for 施工方案、施工组织设计、专项方案、专项吊装方案、清单编方案、模板改方案、规范原文截图、危险源辨识、应急预案、施工进度计划, especially when the user requires the established 冷轧厂 or 条材厂 cover and approval layout, preserved font hierarchy, BOQ-specific methods, editable lifting calculations, exact crane load-chart evidence, downloaded standard-PDF excerpts, generated construction diagrams, and automatic depth plus render verification.
---

# 现场施工方案 Skill

## 文字输出风格：起草与交付前审查

撰写报告、方案、清单说明、表格备注或最终答复前，先读随技能内置的[非防御性写作适配](references/anti-defensive-writing/STYLE.md)。先讲有证据的结论，再说明依据、影响和执行动作；按该指南完成交付前文字审查。

文字优化不得改动用户范围、数量价格、引用依据、真实缺陷、必要暂估条件、模板约束及专业验收要求。采用本地相对引用，不依赖另行安装写作插件。专业复核完成后，可用[只读文字检查器](scripts/review_prose.py)辅助检查重复套话和指定必须保留的事实；它不替代事实与工程审查。

Produce a finished, editable construction scheme rather than an outline or compressed technical note. Preserve the selected plant template's cover, approval page, and typography while building a full, project-specific technical and management body. Substantiate methods with current standards, use visuals in balance with prose, and block delivery when depth or render gates fail.

## Default Contract

- Write in professional Chinese for owner, supervision, safety, construction, and approval review.
- Use `武汉兴达建筑安装工程有限责任公司` when the user does not name another contractor.
- Treat supplied start and finish dates as inclusive; compute and state calendar days accurately.
- Preserve personnel names and staffing from the source scheme. Never invent named personnel.
- Preserve the source/template file. Write to a new project-specific output folder.
- Treat presentation and content depth as separate decisions. A template controls the cover, approval page, contents, heading ladder, fonts, spacing, tables, header, footer, and page numbering; it never justifies a shorter body.
- Select the presentation template in this order: user-supplied template; user-named plant style; project-name match (`冷轧厂` or cold-rolling unit -> `assets/冷轧厂施工方案版式模板.docx`, `条材厂` -> `assets/条材厂施工方案版式模板.docx`); otherwise use the bundled 条材厂 style and state the choice.
- Treat every block from the cover through the last contents entry as a frozen front-matter region. Replace only approved variable text at run level: project name, project number, contractor, date, and supplied signatory names. Preserve fixed labels and all formatting and structure.
- Preserve existing point sizes unless the user explicitly asks to resize text. Change font face without flattening the size hierarchy.
- Preserve the selected template's body face. Where the source does not define one, use `方正仿宋_GB2312`; keep cover, headings, captions, tables, and schedule at their established faces and sizes.
- Default to the `standard` full-depth profile. Use `complex` for multi-discipline, shutdown/tie-in, hazardous media, lifting, confined-space, pressure, or major production-interface work. Use `small` only when the user explicitly asks for a concise scheme and the scope is genuinely single-discipline and low-risk.
- Deliver the DOCX only after automatic content-depth checks and Word/PDF visual inspection both succeed.

## Load Supporting Guidance

Read only the references needed for the task:

- Read `references/template-rules.md` before creating or restructuring a DOCX.
- Read `references/content-checklist.md` before drafting technical chapters.
- Read `references/table-policy.md` before deciding whether any body content belongs in a table.
- Read `references/quality-baseline.md` before choosing the depth profile and before final validation.
- Read `references/standards-and-evidence.md` when standards must be researched, downloaded, or illustrated.
- Read `references/v2-contract.md` before initializing a project or writing a project builder.
- Read `references/lifting-special-plan.md` for any standalone lifting plan or lifting chapter that selects a crane, lifting points, rigging, spreader beam, outriggers, or a crane standing position.
- Read `references/verification.md` before final handoff.

Use the `documents` skill for DOCX work, the `spreadsheets` skill for BOQ work, the `pdf` skill for standard PDFs, and `imagegen` when a generated bitmap schematic materially improves understanding.

## V2 Operating Rule

Use a project contract as the single machine-readable source of truth. Initialize every new working directory with `scripts/init_project.py`, then update `project-contract.json` after inspecting the BOQ, drawings, source scheme, and template. Project builders may import `scripts/docx_builder.py` and `scripts/project_contract.py`; they must never import or dynamically load another project's `build_scheme.py`, `PREVIOUS_UTILS`, `BASE_BUILDER`, or an absolute script path below a prior `work` directory.

## Workflow

### 1. Intake And Protect Sources

1. Resolve every supplied BOQ, drawing, photo, scheme, and template path.
2. Copy or read sources into a dedicated working directory; never overwrite the originals.
3. Identify the project name, contractor, theoretical dates, required output type, fixed personnel list, plant style, and depth profile.
4. Inspect the accepted template before drafting. Record section orientation, margins, cover positions, approval table, contents page, heading faces and sizes, body face and size, paragraph rhythm, table sizes, footer/page number, and schedule layout.
5. Record a short template decision in the work log: `冷轧厂`, `条材厂`, or `user-supplied`, with the source path. If the user says only the font should change, preserve all point sizes and spacing.
6. Identify the first body-heading marker and its occurrence in the template. Use it later with `scripts/audit_front_matter.py`; do not edit front-matter page breaks, tables, margins, section properties, headers, footers, or styles to fix pagination.
7. Run `scripts/init_project.py` and complete `project-contract.json`. Record fixed inclusive dates when supplied; when only duration is known, use the D-day schedule basis and do not invent calendar dates.

### 2. Extract And Interpret The BOQ

Run:

```powershell
python scripts/inspect_boq.py "<boq.xlsx>" --output "<workdir>\boq-analysis.json"
```

Then inspect the extracted rows directly. Determine:

- major disciplines and interfaces;
- demolition, isolation, temporary works, civil, structural, mechanical, piping, electrical, instrumentation, anticorrosion, landscaping, drainage, and commissioning scope as applicable;
- large or hazardous equipment, estimated weights, elevations, restricted access, cranes, aerial work platforms, scaffolds, shutdown windows, and production tie-ins;
- quantities that should appear in the scheme and quantities that require drawing or site confirmation.

Create a project coverage contract before drafting:

- 6-12 project-specific terms taken from the BOQ, drawings, equipment, media, locations, or production interfaces;
- 8-16 major method terms representing the actual work packages;
- all applicable disciplines and high-risk operations;
- legacy project names and irrelevant disciplines that must not remain;
- a BOQ-to-chapter matrix mapping each major scope item to method, quality, safety, resources, schedule, and acceptance content.

Do not convert uncertain quantities into facts. Mark drawing-, manufacturer-, owner-, or site-controlled values for confirmation.

### 3. Build A Project-Specific Technical Logic

Draft construction logic from the actual scope, not generic prose. Use the complete content framework in `references/content-checklist.md`; preserve the selected template's visual hierarchy while mapping all required modules into it. Establish:

1. preconditions, survey, permits, shutdown, LOTO, medium evacuation, cleaning, detection, and handover;
2. demolition and waste handling sequence;
3. civil foundations and embedded parts;
4. equipment, structure, piping, electrical, instrumentation, anticorrosion, and specialist methods;
5. hold points, witness points, inspection records, tests, flushing, loop checks, interlocks, no-load/load commissioning, and production handover;
6. cross-discipline interfaces and acceptance responsibility.

For every major method, write preconditions, resources, sequence, verified technical parameters, quality checks, safety controls, records, and acceptance as prose under numbered headings. Apply the table-admission test in `references/table-policy.md`: use a table only when readers must compare repeated records across shared fields. Do not count copied standards, schedules, approval tables, or BOQ tables as technical depth.

### 3A. Activate The Lifting Special-Plan Mode

For a standalone lifting plan, major lifting chapter, crane selection, or lifting calculation, initialize with `--special-plan lifting`, force the `complex` profile, and follow `references/lifting-special-plan.md`. Complete `lifting-data-register.json` before drafting calculations. Treat each lifted component and each governing crane configuration as a separate traceable record.

Do not use web search to invent project-specific component weight, center of gravity, ground bearing capacity, lifting-point geometry, or standing coordinates. Web research may obtain current official standards and the exact manufacturer's manual or load chart. If owner, drawing, fabrication, weighing, geotechnical, survey, or site data remain unavailable, list the missing inputs and block an execution-ready issue.

Use `add_omml_equation()` from `scripts/docx_builder.py` for the calculation appendix. Every formula must be an editable Word OMML equation; Latin letters, Greek symbols, numbers, units, and mathematical operators use Times New Roman. Show input source, symbol definition, formula, numeric substitution, result, acceptance criterion, conclusion, and independent checker. Plain-text formula substitutes and formula screenshots do not satisfy this requirement.

The lifting audit requires component actual weight and center-of-gravity evidence; exact crane manufacturer, model, counterweight, boom, outrigger, radius, rated capacity, and original load-chart page; complete checks for design load, crane capacity, rigging, lifting points, spreader beam when used, maximum outrigger reaction, ground bearing, standing setout, hook height and clearance, trial lift, and wind control; plus verified lifting-point, standing, outrigger, and rigging figures. `scripts/validate_scheme.py` automatically runs `scripts/audit_lifting_plan.py` and blocks delivery when this evidence is missing.

### 4. Research Standards And Create Evidence

1. Search the web because standard status can change.
2. Prefer official national, ministry, industry, or publisher sources for current-status verification.
3. Download the standard PDF when legally and technically available.
4. Immediately archive every newly downloaded standard in the versioned library:

```powershell
python scripts/archive_standard.py "<downloaded-standard.pdf>" --number "<standard-number>" --title "<title>" --source-kind official --source-url "<url>" --project "<project-name>"
```

5. Use the archived PDF under `standards/library`, not the temporary download, for all later work.
6. Inspect the actual PDF and identify clauses that directly support the method or control measure.
7. Render those PDF pages with `scripts/render_standard_pages.py`.
8. Insert the rendered PDF page, not a browser or website screenshot.
9. Add a concise caption containing standard number, topic, PDF page, source filename, and retrieval date.
10. Copy the applicable archived standards into the final project folder under `引用规范原文` with an index.

Limit excerpts to useful evidence. Explain the project's application before or after every excerpt. Keep at least 4,000 narrative characters per standard excerpt and avoid alternating text pages with full-page screenshots. See `references/standards-and-evidence.md`.

Never silently replace a library PDF. The archive script stores by hash, preserves conflicting editions, and updates `standards/index.json`. Treat new or changed library files as pending Git changes; report them before any future commit or push.

### 5. Create Visuals Only Where They Help

Use drawing crops, site photographs, or generated technical schematics for complex routing, lifting, PPH fusion, process flow, temporary protection, or emergency evacuation. Prefer 2-6 useful technical figures for a normal full scheme when the scope supports them; never add decorative filler.

For generated visuals:

- use `imagegen`;
- use a clean white engineering background;
- omit fabricated dimensions and text labels unless verified;
- remove empty callout boxes, placeholder bars, logos, and watermarks;
- caption the image as a schematic and state that drawings/site conditions govern.

### 6. Assemble The DOCX

1. Start from the supplied or bundled template.
2. Import stable helpers from `scripts/docx_builder.py`. Keep project facts in `project-contract.json`; do not inherit code or constants from a previous project builder.
3. Use run-level text replacement. Do not use `paragraph.text = ...` on retained template paragraphs.
4. Do not run a whole-document sanitizer or style normalizer.
5. Keep the cover, approval page, contents page, page numbering, and chapter hierarchy. Do not delete page breaks, force approval-table row heights, rewrite fixed approval labels, change cell margins, or modify front-matter paragraph alignment.
6. Put the detailed schedule in a landscape final section.
7. Use stable table widths, repeating header rows, non-splitting rows, restrained header shading, and thin borders.
8. Keep screenshots legible and captions attached to their figures.
9. Avoid orphan headings, nearly blank pages, stretched images, and tables that dominate the narrative.
10. Rebuild the body from the complete framework. Do not retain an old project's body merely because its cover and approval page are being reused.
11. Keep normal body tables to a target of 6-10 beyond fixed front matter; this is a ceiling target, not a quota. Combine related resource or inspection data instead of creating many small tables. Convert prose-heavy cells into numbered paragraphs.

### 7. Build The Schedule

1. Compute inclusive duration from the user-supplied dates.
2. Decompose work according to actual dependencies and production constraints.
3. Show overlapping professional work only when interfaces permit it.
4. Include preparation, isolation, demolition, civil work, fabrication/procurement, installation, electrical/instrumentation, tests, commissioning, restoration, and handover as applicable.
5. State that approved weekly plans may refine the bars without breaking the contractual completion date.
6. Render a readable landscape Gantt or schedule table with exact start/finish dates.

### 8. Verify Before Delivery

1. Run the V2 gate from the skill root:

```powershell
python scripts/validate_scheme.py "<workdir>\project-contract.json"
```

2. Treat any contract, lifting-data, Word-equation, Word export, front-matter, body-font, required-term, depth, schedule, render, DOCX-integrity, or standard-library failure as a delivery blocker.
3. Inspect every generated contact sheet and all critical pages at full size. Automatic pixel checks supplement visual review; they do not replace it.
4. Review `qa/validation-summary.json` and confirm that approved front-matter changes and any render exception are narrowly scoped.
5. Confirm all output files and required archived standard PDFs exist.
6. Iterate until `references/verification.md` passes. Never lower a threshold to make an incomplete document pass; improve the document or record a user-approved exception.

## Example Invocations

- `用 $site-construction-scheme 按冷轧厂版式，根据这个清单做完整施工方案，工期8月7日至12月31日。`
- `用 $site-construction-scheme 按条材厂封面审批表和字体版式，根据这个清单做完整施工方案。`
- `用 $site-construction-scheme 修改这份别人编的方案，人员不要改，补规范原文截图、危险源和应急预案。`
- `用 $site-construction-scheme 把专项吊装方案统一格式后并入原施工方案，正文只改为方正仿宋_GB2312，字号不要变。`
- `用 $site-construction-scheme 根据构件图、现场总图和吊机资料编制专项吊装方案，逐件核对实重与重心，完成吊点、平衡梁、支腿反力、地基承载力和站位放样计算。`

## Handoff

Report the final DOCX path, project-contract path, validation-summary path, selected template style, depth profile, page count, narrative character count, chapter and method counts, image/table balance, schedule basis and inclusive duration, standards archived, diagrams inserted, and validation performed. For lifting plans, also report the lifting-data-register path, exact crane model/configuration, governing component and design load, OMML equation count, and lifting audit result. Mention unresolved owner/drawing/site confirmations plainly. Do not present a draft or failed validation gate as a completed deliverable.
