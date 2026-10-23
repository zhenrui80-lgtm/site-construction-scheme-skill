# Template contract

## Select the format profile first

New construction schemes default to `latest-20261023`; read [latest-format.md](latest-format.md) and verify `assets/latest-template-manifest.json`. Its body is 宋体 12 pt / 1.5 lines, with 15 content topics and six appendix forms. Its paragraph counts are expandable and its appendix tables retain the source schema. Clone only approved body/heading/caption/data-row patterns; audit all changed structures. The latest blank template is an explicit one-time baseline migration authorized by the user's new format request, not a license to normalize finished documents.

The remaining fixed-index contract below applies **only to `legacy-transport`**. Never apply those counts, offsets or figure exemptions to the new template. A newer project-specific tender/user instruction wins over either profile; record the resulting contract and re-distill new formats.

## Legacy transport template

## Preservation rules

The supplied template is a contract, not a styling suggestion.

- Preserve the original cover, fonts, font sizes, paragraph properties, line spacing, indentation, tabs, leaders, numbering, table properties, section breaks, margins, headers/footers, and TOC construction.
- Use run-level replacement: change `run.text` inside an existing paragraph. Do not assign `paragraph.text` or `cell.text` for replacements because doing so recreates runs and destroys formatting.
- Do not run a format calibrator, normalizer, style cleanup, or document-wide font/spacing/indent pass.
- Do not recreate all paragraphs or tables. Do not repair cosmetic inconsistencies unless the user expressly requests a targeted repair.
- Preserve the source template unmodified. Generate to a new output path.

## Known fixed-template fingerprint

For `assets/template_transport_office_scheme.docx`, the expected SHA-256 is stored in `assets/template.sha256`. The observed baseline structure is:

- 2,026 body paragraphs;
- 20 tables;
- 105 sections;
- Chapter 1 TOC text slots at zero-based paragraph indices 26 through 39;
- Chapter 1 body range at zero-based paragraph indices 208 through 297;
- intended Chapter 1 figure anchor offsets 2, 39, 50, 62, 71, and 80 within that body range, with each immediately following paragraph reserved for its caption;
- intended Chapter 1 heading offsets 0, 4, 10, 16, 22, 28, 34, 44, 57, 66, 76, 84, 86, and 88.

These offsets are valid only when the hash matches. If it differs, inspect the new template structure and derive a new allowed-slot map. Never use stale offsets against a changed template.

## Safe replacement pattern

For each target paragraph:

1. Locate runs containing visible text while preserving tab/leader runs.
2. Put the replacement in the first eligible text run.
3. Clear only the text of the remaining eligible text runs.
4. Do not change run properties or paragraph properties.
5. Verify the paragraph still has the expected tabs, style, alignment, spacing, and indentation.

For a longer replacement, prefer approved empty/placeholder paragraphs in the same chapter. Do not expand by inserting uncontrolled paragraphs that shift the template hierarchy.

## Figures

- Insert into a designated placeholder paragraph without replacing that paragraph wholesale.
- Preserve paragraph properties. Set only the picture run and caption text needed for the approved slot.
- Keep the image within page width and verify that it does not create a blank page or clip headers/footers.
- Source drawing extracts are preferred where available. Generated conceptual figures must be labelled as schematic and must not imply approved dimensions.

## Structure audit

Outside explicitly allowed text/figure slots, compare template and output for:

- paragraph count and order;
- paragraph properties and style IDs;
- run count and run properties;
- table count/shape/properties;
- section count/properties;
- header/footer relationships;
- package integrity and relationships.

Treat any unexplained drift as a failure.

Run `scripts/audit_template_fidelity.py <template.docx> <output.docx>`. Default figure/paired-caption exemptions apply only when the legacy SHA matches. Other templates receive no automatic old-index exemptions; add only individually inspected explicit `--allow-body-paragraph` locations.
