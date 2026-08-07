# Conversation Requirements Archive

This archive records the reusable requirements and source artifacts from the construction-scheme workflow that led to this Skill. Binary files are listed in `manifest.json`.

## Chronological Requirements

1. Prepare a professional 10,000-character-class landscape-improvement construction scheme from BOQ and CAD/PDF drawings for `武汉兴达建筑安装工程有限公司`, including garden structures, water supply and drainage, electrical work, landscaping, a water feature, and a fixed progress plan.
2. Improve the Word layout and insert useful images cropped from drawings.
3. Use an accepted WPS construction-organization design as the formatting template and correct chaotic layout.
4. Reuse the template for a new 40-day corrosion-remediation project.
5. Prepare a new 武钢日铁 electrical-routing scheme: remove noncompliant wiring, route new cables in trays, and run conduit down columns at about 5 m elevation; use the supplied ladder-replacement scheme and site photos as references.
6. Prepare a 100-day coke-plant waste-gas collection and environmental-remediation scheme from a new BOQ.
7. Learn and reproduce the formal 条材厂 scheme format and search authoritative standards for accurate citations.
8. Prepare the quality-center canteen gas-to-electric scheme for 2026-08-10 through 2026-12-31. Download applicable standards, insert standard evidence, and generate useful schematics.
9. Correct the header/approval layout to match the 条材厂 template. Add emergency response, hazard identification, and more narrative explanation. Avoid uncomfortable website screenshots and excessive tables.
10. Revise an underdeveloped section-steel descaling-system scheme without changing personnel. Download standards from authoritative sources and insert screenshots from the actual standard PDF pages.
11. Store downloaded standards in the project library. Extract and insert clauses that the scheme actually cites. Improve the personnel list layout.
12. Reformat and merge a high-pressure water-tank lifting scheme into the original system-upgrade scheme. Use `方正仿宋_GB2312` for body text while preserving existing point sizes.
13. Prepare the `101机组酸循环系统和酸雾系统综合改造` scheme from `报价清单.xlsx` for 2026-08-07 through 2026-12-31, including downloaded standard-PDF excerpts and generated technical diagrams.
14. Turn the complete workflow into a reusable Skill in a private Git repository. Archive this conversation's meaningful source, final, script, visual, and standards artifacts. Automatically save every newly downloaded standard into the versioned library.

## Stable Decisions

- Default contractor: `武汉兴达建筑安装工程有限责任公司`.
- Default template: the sanitized 条材厂-style template bundled under `assets`.
- Body face: `方正仿宋_GB2312`.
- Preserve the template size hierarchy; changing font does not authorize changing size.
- Preserve source personnel names and configuration.
- Keep cover, owner approval page, contents, ten main chapters, page numbering, and a landscape final schedule.
- Use narrative explanation as the main content; use tables to organize supporting data.
- Include project-specific hazard identification and emergency response.
- Verify standard status online, download actual PDFs, and insert images rendered from the PDF pages. Do not insert standards-website screenshots.
- Archive standards by SHA-256 and retain multiple editions or conflicting scans.
- Use generated schematics only where they improve technical understanding. Remove fabricated labels, empty callout boxes, logos, and watermarks.
- Render the final DOCX to PDF and inspect all pages before delivery.

## Archive Scope

Included:

- available user-supplied BOQ, drawing archive, template, scheme, lifting-scheme, and reference documents mentioned in this conversation;
- completed and intermediate meaningful scheme outputs from the corresponding work folders;
- reusable generation, extraction, formatting, and audit scripts;
- selected diagrams, drawing crops, standard excerpts, and current QA PDF;
- all readable standard PDFs recovered from the conversation work libraries.

Excluded because they are reproducible or unsuitable for source control:

- Word/PDF page-render PNG batches and contact-sheet caches;
- browser profiles and browser caches;
- `node_modules` and downloaded package caches;
- duplicate extracted drawing files when the original archive is preserved;
- duplicate binaries already represented once by SHA-256.

Two original WeChat temporary JPG paths referenced in the conversation were no longer present on disk when this archive was created. Their missing source-path status is recorded in `manifest.json`; equivalent recovered copies are included as `work-products/04-wugang-wiring/site_01.jpg` and `site_02.jpg`.
