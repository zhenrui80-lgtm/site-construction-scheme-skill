---
name: industrial-construction-scheme
description: Create and revise professional Chinese metallurgical or industrial construction-scheme DOCX deliverables from BOQ spreadsheets, drawings, photos, existing schemes, and fixed schedules. Use for 施工方案、施工组织设计、专项方案、清单编方案、模板改方案、规范原文截图、危险源辨识、应急预案、施工进度计划, especially when the user requires the established 条材厂-style approval layout, 方正仿宋_GB2312 body text, downloaded standard-PDF excerpts instead of website screenshots, generated construction diagrams, and render-verified Word output.
---

# Industrial Construction Scheme

Produce a finished, editable construction scheme rather than an outline. Analyze the supplied artifacts, preserve the accepted template, substantiate methods with current standards, add only useful visuals, and verify the final DOCX page by page.

## Default Contract

- Write in professional Chinese for owner, supervision, safety, construction, and approval review.
- Use `武汉兴达建筑安装工程有限责任公司` when the user does not name another contractor.
- Treat supplied start and finish dates as inclusive; compute and state calendar days accurately.
- Preserve personnel names and staffing from the source scheme. Never invent named personnel.
- Preserve the source/template file. Write to a new project-specific output folder.
- When no template is supplied, use `assets/条材厂施工方案版式模板.docx`.
- Preserve existing point sizes unless the user explicitly asks to resize text. Change font face without flattening the size hierarchy.
- Default body face to `方正仿宋_GB2312`; keep cover, headings, captions, tables, and schedule at their established sizes.
- Deliver the DOCX only after Word/PDF rendering and visual inspection succeed.

## Load Supporting Guidance

Read only the references needed for the task:

- Read `references/template-rules.md` before creating or restructuring a DOCX.
- Read `references/content-checklist.md` before drafting technical chapters.
- Read `references/standards-and-evidence.md` when standards must be researched, downloaded, or illustrated.
- Read `references/verification.md` before final handoff.

Use the `documents` skill for DOCX work, the `spreadsheets` skill for BOQ work, the `pdf` skill for standard PDFs, and `imagegen` when a generated bitmap schematic materially improves understanding.

## Workflow

### 1. Intake And Protect Sources

1. Resolve every supplied BOQ, drawing, photo, scheme, and template path.
2. Copy or read sources into a dedicated working directory; never overwrite the originals.
3. Identify the project name, contractor, theoretical dates, required output type, and any fixed personnel list.
4. Inspect the accepted template before drafting. Record section orientation, margins, cover positions, approval table, contents page, heading sizes, body size, table sizes, footer/page number, and schedule layout.
5. If the user says only the font should change, preserve all point sizes and spacing.

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

Do not convert uncertain quantities into facts. Mark drawing-, manufacturer-, owner-, or site-controlled values for confirmation.

### 3. Build A Project-Specific Technical Logic

Draft construction logic from the actual scope, not generic prose. Establish:

1. preconditions, survey, permits, shutdown, LOTO, medium evacuation, cleaning, detection, and handover;
2. demolition and waste handling sequence;
3. civil foundations and embedded parts;
4. equipment, structure, piping, electrical, instrumentation, anticorrosion, and specialist methods;
5. hold points, witness points, inspection records, tests, flushing, loop checks, interlocks, no-load/load commissioning, and production handover;
6. cross-discipline interfaces and acceptance responsibility.

Use the chapter and depth rules in `references/content-checklist.md`. Narrative explanation should outweigh tables; tables should organize quantities, responsibilities, hazards, resources, checks, and schedules.

### 4. Research Standards And Create Evidence

1. Search the web because standard status can change.
2. Prefer official national, ministry, industry, or publisher sources for current-status verification.
3. Download the standard PDF when legally and technically available.
4. Immediately archive every newly downloaded standard in the versioned library:

```powershell
python scripts/archive_standard.py "<downloaded-standard.pdf>" --number "<standard-number>" --title "<title>" --source-url "<url>" --project "<project-name>"
```

5. Use the archived PDF under `standards/library`, not the temporary download, for all later work.
6. Inspect the actual PDF and identify clauses that directly support the method or control measure.
7. Render those PDF pages with `scripts/render_standard_pages.py`.
8. Insert the rendered PDF page, not a browser or website screenshot.
9. Add a concise caption containing standard number, topic, PDF page, source filename, and retrieval date.
10. Copy the applicable archived standards into the final project folder under `引用规范原文` with an index.

Limit excerpts to useful evidence. Avoid turning the document into alternating blank text pages and full-page screenshots. See `references/standards-and-evidence.md`.

Never silently replace a library PDF. The archive script stores by hash, preserves conflicting editions, and updates `standards/index.json`. Treat new or changed library files as pending Git changes; report them before any future commit or push.

### 5. Create Visuals Only Where They Help

Use drawing crops, site photographs, or generated technical schematics for complex routing, lifting, PPH fusion, process flow, temporary protection, or emergency evacuation.

For generated visuals:

- use `imagegen`;
- use a clean white engineering background;
- omit fabricated dimensions and text labels unless verified;
- remove empty callout boxes, placeholder bars, logos, and watermarks;
- caption the image as a schematic and state that drawings/site conditions govern.

### 6. Assemble The DOCX

1. Start from the supplied or bundled template.
2. Use run-level text replacement. Do not use `paragraph.text = ...` on retained template paragraphs.
3. Do not run a whole-document sanitizer or style normalizer.
4. Keep the cover, approval page, contents page, page numbering, and chapter hierarchy.
5. Put the detailed schedule in a landscape final section.
6. Use stable table widths, repeating header rows, non-splitting rows, restrained header shading, and thin borders.
7. Keep screenshots legible and captions attached to their figures.
8. Avoid orphan headings, nearly blank pages, stretched images, and tables that dominate the narrative.

### 7. Build The Schedule

1. Compute inclusive duration from the user-supplied dates.
2. Decompose work according to actual dependencies and production constraints.
3. Show overlapping professional work only when interfaces permit it.
4. Include preparation, isolation, demolition, civil work, fabrication/procurement, installation, electrical/instrumentation, tests, commissioning, restoration, and handover as applicable.
5. State that approved weekly plans may refine the bars without breaking the contractual completion date.
6. Render a readable landscape Gantt or schedule table with exact start/finish dates.

### 8. Verify Before Delivery

1. Update contents, fields, page numbers, and pagination in Word.
2. Export the DOCX to PDF. Use `scripts/export_docx_word.ps1` on Windows when Word COM is available.
3. Render every PDF page to PNG and inspect contact sheets plus critical pages at full size.
4. Run `scripts/audit_docx.py` with required project terms, forbidden legacy terms, and expected font.
5. Check the DOCX ZIP package for corruption.
6. Confirm all output files and archived standard PDFs exist.
7. Iterate until `references/verification.md` passes.

## Example Invocations

- `用 $industrial-construction-scheme 按我们之前的模板，根据这个清单做施工方案，工期8月7日至12月31日。`
- `用 $industrial-construction-scheme 修改这份别人编的方案，人员不要改，补规范原文截图、危险源和应急预案。`
- `用 $industrial-construction-scheme 把专项吊装方案统一格式后并入原施工方案，正文只改为方正仿宋_GB2312，字号不要变。`

## Handoff

Report the final DOCX path, page count, inclusive duration, standards archived, diagrams inserted, and validation performed. Mention unresolved owner/drawing/site confirmations plainly. Do not present a draft as a completed deliverable.
