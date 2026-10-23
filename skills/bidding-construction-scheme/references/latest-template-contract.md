# New blank-template edit contract

Read with `latest-format.md`; fingerprints are in `assets/latest-template-manifest.json`.

## Source and intentional migration

The inspected 8-page source had 49 body paragraphs, four blank tables, seven sections and an unintended blank page 2. It is a format/requirements reference, not a completed scheme. The preserved public reference removes only author and custom metadata. The derived template reorganizes this into a neutral cover, native TOC, 15 topic slots and six in-chapter appendices, removing the reference instruction page and redundant section breaks. These explicit baseline changes are authorized by the user's request to adopt the latest format. No prior project facts or bidder identity are embedded.

Source A4 geometry is retained: 11906 × 16838 twips, top/bottom 1440, left 1797, right 1644, header 851 and footer 992; one column. The original four table XML trees are copied verbatim, including widths, row heights, cell paragraphs and merged stage header. Annex titles and body placeholders are new role-based paragraphs. A centered dynamic PAGE field is added for pagination; no company header/logo is introduced.

## Editable roles

- `BiddingCover`: neutral project name, document title and duration; 宋体16 pt bold centered, 1.5 lines. Add bidder/date only when supplied and permitted by tender identity rules.
- `BiddingBody`: 宋体12 pt, 1.5 lines, 24 pt first-line indent, zero paragraph before/after spacing. Clone this paragraph for prose; preserve run/paragraph properties.
- `BiddingHeading1`: 宋体16 pt bold, 1.5 lines, outline 0, keep-with-next. Map all 15 required topics.
- `BiddingHeading2`: 宋体14 pt bold, 1.5 lines, outline 1, keep-with-next. Clone for BOQ methods/subsections.
- `BiddingAppendix`: 宋体12 pt bold, 1.5 lines, outline 1, keep-with-next. Six source-title anchors; table/figure content follows.
- `BiddingCaption`: 宋体10.5 pt centered, 1.5 lines. Use for figure captions/units, not prose disguised as captions.
- `BiddingTOC` and `TOC 1/TOC 2`: native field/result roles; preserve tabs/leaders and numeric-run semantics during field updates. Do not replace a TOC with `p.text`.
- `BiddingFooter`: 宋体10.5 pt centered dynamic PAGE. Retain field wrappers; never type one static page number throughout the document.

## Forms and expansion

Preserve four source table schemas: 10 columns (equipment), 9 columns (instruments), 8 columns (labor with last 7 header cells merged), 4 columns (temporary land). Fill content in existing runs, treating merged cells as one underlying XML cell. Clone or remove **blank data rows only** as required, retaining grid, borders and headers. Record changes; title count alone is not a filled-form check. Table typography follows source rather than full-document 12 pt rewriting. If exact source row heights cause clipping, first shorten factual cell text or add an authorized row; any height change must be scoped, recorded and rendered, never applied as a global cleanup.

Replace the two figure placeholders with schedule/network and site-layout evidence. Insert Chapter 1 figures at newly documented body/caption clones. Keep images within the available text width and use correct captions/provenance. All placeholders `【填写：…】`, `【项目名称】`, `【工期】` must disappear from a completed scheme.

## Gate boundaries

Preparation rejects any source whose SHA differs from the inspected conversion. A fresh WPS/Word conversion can change metadata bytes; only after inspecting the complete new conversion may the operator supply `--expected-source-sha256`. Shape checks alone do not prove an input is an approved blank reference. Never use this preparation command on a completed project.

The skeleton intentionally has empty forms and instructional placeholder text, so template QA is not release QA. Render a working copy, update fields in WPS/Word and inspect every page; never let a render helper overwrite the published template silently. Recheck hashes after any intentional baseline change. A populated project additionally needs engineering review, BOQ coverage, all five Excel sheets, final-day validation and the final rendered-format audit.
