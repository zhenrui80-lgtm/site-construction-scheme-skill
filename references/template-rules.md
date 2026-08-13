# Template And Layout Rules

## Template Selection

Presentation and content depth are independent. Use a template for visual identity and front matter only; always build the complete body required by `content-checklist.md`.

Select in this order:

1. a user-supplied template or explicit instruction;
2. `assets/冷轧厂施工方案版式模板.docx` when the user requests 冷轧厂 style or the project belongs to 冷轧厂/cold-rolling units;
3. `assets/条材厂施工方案版式模板.docx` when the user requests 条材厂 style or the project belongs to 条材厂;
4. the 条材厂 template as the disclosed fallback for other plants without a supplied template.

Do not blend the two styles. Record the selected template path in the task work log.

## Front Matter Fidelity

For either bundled style, preserve:

1. cover geometry, title positions, plant/project fields, contractor line, and date placement;
2. owner approval page structure, labels, signature cells, row heights, borders, and reserved handwriting space;
3. contents-page hierarchy and page-number behavior;
4. header, footer, page numbering, margins, and section orientation;
5. the template's font family, point-size ladder, bold rules, indentation, alignment, line spacing, and paragraph spacing.

Replace text at run level. Rebuild only the body region. Remove all legacy project text, media, tables, bookmarks, and stale fields from the replaced body. Keep the main document portrait and the detailed final schedule landscape.

Treat the cover, approval page, and contents as a frozen region. Approved changes are limited to project name, project number, date, supplied signatory names after unchanged labels, contents text/page numbers required by the new body, and a different contractor only when the user explicitly names one. Preserve all fixed approval labels and the template contractor otherwise.

Never perform these operations in the frozen region:

- delete, add, or move a page break or section break;
- force approval-table row heights or column widths;
- change merged cells, borders, shading, cell margins, vertical alignment, or table positioning;
- clear and rebuild approval cells;
- change paragraph alignment, indentation, line spacing, before/after spacing, or pagination controls;
- apply a body font or body table formatter to front-matter paragraphs or tables;
- replace an inherited font or size with a hard-coded fallback;
- add, remove, or restyle headers and footers;
- modify `Normal`, heading, table, or other shared styles after loading the template.

If a project title wraps badly, preserve formatting and adjust only the title text layout within the existing runs. Do not reshape the table or page. If the template itself contains a genuine defect, create a separately approved template revision and validate it before using it; do not repair the template silently during project generation.

## Body Structure

The original 条材厂 ten-chapter sequence is a presentation convention, not a content ceiling. A full scheme must cover all fifteen modules in `content-checklist.md`. It may:

- use the fifteen-module sequence directly; or
- preserve a user-mandated top-level sequence and map the missing management modules into clearly numbered chapters, subchapters, or appendices.

Do not omit construction preparation/layout, schedule guarantees, seasonal controls, finished-product protection, coordination, commissioning/handover, or warranty merely because the source template lacks those headings.

## Typography

- Preserve the selected template's body face. Use `方正仿宋_GB2312` only when the source has no reliable body definition.
- Preserve the template point sizes. Do not make every run the same size.
- Keep cover title, contents title, chapter headings, body, captions, source notes, and tables as distinct size levels.
- Use bold for chapter and subsection headings according to the template.
- Keep body paragraphs justified with the template's first-line indent and line spacing.
- Use smaller but readable caption, source-note, and table text.

When changing font face, set ASCII, high ANSI, East Asian, and complex-script font attributes at run level. Do not overwrite paragraph text or inherited size values.

## Tables

- Apply `table-policy.md` before creating a body table.
- Keep fixed template approval/signature tables.
- Use body tables only for genuinely comparable repeated records such as a hazard matrix, compact resource plan, inspection/test plan, and Gantt schedule.
- Use prose for project scope interpretation, difficulties, reasoning, sequence, construction methods, quality and safety logic, standards interpretation, emergency procedures, interfaces, coordination, protection, and acceptance logic.
- Keep tables full-width but not crowded; use stable column widths.
- Repeat header rows and prevent row splitting where possible.
- Apply restrained pale blue or gray header shading and thin gray/black borders.
- Avoid nested cards or decorative table layouts.
- Do not create a table merely to make a short section look complete. Do not place multi-paragraph explanations inside cells.

## Figures And Standard Excerpts

- Center figures and keep aspect ratio.
- Keep the figure, caption, and source note together.
- Use a readable width; do not stretch low-resolution images.
- Do not leave generated empty labels or placeholder boxes in diagrams.
- Avoid forcing every figure onto a new page if that produces a nearly empty preceding page.
- If a full-page standard excerpt is needed, place related interpretation immediately before or after it.

## Safe DOCX Editing

- Preserve the original template.
- Prefer run-level replacement for retained cover, approval, and contents text.
- Remove and rebuild only the intended body region when the old project content must be replaced.
- Do not use `p.text = ...` on retained formatted paragraphs.
- Do not apply a global paragraph normalization pass.
- Do not mutate shared styles after loading a template. Apply body formatting only to newly created body runs and paragraphs.
- If the target is open in Word/WPS, save to a new filename rather than overwriting.
