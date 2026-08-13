# Final Verification

## Content

- Project name, contractor, dates, inclusive duration, scope, and quantities agree with inputs.
- No legacy project name, old discipline, old equipment, or old schedule remains.
- Personnel names match the source and were not invented.
- Major BOQ disciplines and high-risk operations appear in methods, quality, safety, resources, and schedule.
- Hazard identification and emergency response are project-specific.
- Standard citations are current-status checked and excerpts come from downloaded PDFs.
- Generated diagrams are useful, clean, and identified as schematic.
- All fifteen content modules are present, even when a user-mandated top-level sequence groups them differently.
- Every major BOQ work package appears in method, quality, safety, resources, schedule, and acceptance content.
- Every major method includes preconditions, resources, sequence, parameters, quality, safety, records, and acceptance.
- Every body table passes the table-admission test; explanatory methods and management measures remain prose.

## Formatting

- Cover, owner approval page, contents, chapter order, footer, and page numbers are intact.
- Frozen front matter matches the exact selected template for block order, page/section breaks, approval-table geometry, merged cells, borders, cell margins, paragraph formatting, run formatting, referenced styles, headers, and footers. Only approved variable text differs.
- Body face is `方正仿宋_GB2312` where required.
- Existing point-size hierarchy is preserved.
- Paragraphs are readable and justified; headings are not stranded.
- Tables fit inside margins, repeat headers, and do not split incoherently.
- Normal and complex schemes keep explanatory text at or above 82%, no more than 4 tables and 35 table rows per 10,000 body characters, and no unexplained prose-heavy tables.
- Figures preserve aspect ratio; captions and source notes stay with figures.
- No unexpected blank pages, clipping, overlap, stretched images, or tiny unreadable screenshots.
- Final schedule is landscape and readable at normal zoom/print size.

## Technical Validation

- DOCX ZIP integrity test returns no bad member.
- DOCX reopens after Word field update and save.
- PDF export page count is plausible.
- Every rendered page is nonblank and visually inspected through contact sheets.
- Cover, approval page, contents, representative text pages, every standard excerpt, every schematic, hazard table, emergency section, resource tables, and schedule receive full-size inspection.
- Required-term and forbidden-term audit passes.
- `scripts/audit_scheme_depth.py` passes the selected profile with the task's project terms, method terms, and legacy terms.
- The standard profile is the default; a small profile has a recorded user request and a genuinely low-risk, single-discipline basis.
- Final project folder contains the DOCX and the `引用规范原文` library when standards were requested.

Run both deterministic audits:

```powershell
python scripts/audit_front_matter.py "<selected-template.docx>" "<final.docx>" --body-start "<first-body-heading>" --occurrence 2 --output "<workdir>\front-matter-audit.json"
python scripts/audit_docx.py "<final.docx>" --required "<project-name>" --forbidden "<legacy-name>" --output "<workdir>\docx-audit.json"
python scripts/audit_scheme_depth.py "<final.docx>" --profile standard --project-term "<term-1>" --project-term "<term-2>" --method-term "<method-1>" --method-term "<method-2>" --legacy-term "<legacy-name>" --output "<workdir>\depth-audit.json"
```

Supply the full 6-12 project terms and 8-16 method terms even though the abbreviated example shows fewer. A nonzero exit code blocks delivery.

The front-matter audit allows project-name, project-number, date, signatory values after unchanged `编制/审核/审批/批准` labels, and contents fields. Contractor text and approval labels remain frozen by default. When the user explicitly supplies another contractor or a different template has another verified variable field, add a narrowly scoped `--allow-text-pattern`. Never use a broad pattern that permits fixed approval labels.

## Handoff

State:

- final DOCX path;
- selected `冷轧厂`, `条材厂`, or user-supplied template and depth profile;
- frozen-front-matter audit result and approved variable text changes;
- page count;
- narrative characters, major chapters, heading candidates, substantive management topics, project-specific coverage, table-text ratio, table and row density, prose-heavy tables, and image-text ratio;
- inclusive duration;
- number of standard excerpts and diagrams;
- archived standards folder;
- checks performed;
- unresolved site/drawing/owner confirmations.
