# Standards And Evidence Workflow

## Selection

Select standards from the actual disciplines and hazards. Typical categories include:

- construction quality acceptance and unified acceptance;
- industrial equipment and piping;
- electrical installations, cable lines, grounding, and fire sealing;
- instrumentation and automation;
- building anticorrosion and coating;
- steel structures, welding, concrete, and foundations;
- hazardous chemical special operations;
- lifting machinery and lifting operations;
- work at height, temporary electricity, confined space, hot work, and demolition;
- occupational health, environmental protection, and fire protection;
- owner or metallurgical plant management rules.

Do not cite a standard only because it is common. Link each standard to a scope item, method, acceptance requirement, or risk control.

## Current-Status Verification

Standards are time-sensitive. Verify the standard number, title, edition, implementation status, replacement relationship, and mandatory/general nature using current official or authoritative sources. Record the retrieval date.

If the official page only verifies status but does not supply a downloadable PDF, use the status page as verification and obtain the lawful PDF from an authorized source when available. Never represent an unofficial scan as an official download.

## PDF Evidence

1. Download the full standard PDF to a temporary work location.
2. Run `scripts/archive_standard.py` immediately with standard number, title, source URL, retrieval date, and project name.
3. Confirm that `standards/index.json` contains the SHA-256, page count, archived filename, source, and project usage.
4. Verify the archived file opens, has nonzero pages, and contains the cited clause.
5. Render the actual clause page at readable resolution.
6. Crop only margins if necessary; do not remove the standard number, clause number, table number, or context needed to understand the excerpt.
7. Insert a caption such as `图X  GB 30871-2022 介质处理与能量隔离原文摘录`.
8. Add a source note with archived PDF filename, PDF page, and retrieval date.
9. Follow the image with project-specific interpretation.

Do not use a screenshot of a search result, standards website, browser PDF viewer controls, or paywall page.

## Evidence Balance

- Prefer 4-8 high-value excerpts for a normal scheme.
- Add more only for a technically broad or high-risk project.
- Keep each excerpt close to the relevant method or control.
- Avoid redundant pages from the same standard.
- Use full-page screenshots only when the clause would otherwise be unreadable.
- Archive the downloaded PDFs in `引用规范原文` and add an index mapping standards to scheme chapters and figures.

## Versioned Standard Library

- Store canonical files under `standards/library`.
- Store machine-readable metadata in `standards/index.json`.
- Deduplicate by SHA-256, not only by filename.
- Preserve multiple editions or conflicting scans instead of overwriting.
- Use Git LFS for standard PDFs.
- Record every project that uses a standard.
- Save new standards locally during the task. Do not auto-commit or auto-push without reviewing the pending Git scope.

## Generated Schematics

Generated diagrams supplement evidence; they do not replace drawings or standards. Use them for process relationships, fusion sequence, lifting zones, routing concepts, temporary supports, or emergency flow. Clearly label them as schematic and state that approved drawings, manufacturer documents, and site conditions govern.
