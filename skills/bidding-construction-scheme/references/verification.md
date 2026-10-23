# Delivery verification

All automated checks are necessary but not sufficient. Record results in QA artifacts and inspect the rendered output.

## Latest-format release gate

- Verify `latest-20261023` manifest and source roles. Never compare its variable-length output against legacy paragraph/section counts.
- `audit_latest_format.py` checks actual final PDF count (target <=200), effective body宋体12pt/1.5, chapter/appendix titles, form schemas and placeholders. Independently review completeness: title presence is not proof of substantive content or a usable diagram.
- Count the whole PDF unless the tender states otherwise. Source wording is a recommendation/principle, not automatic disqualification. Over-limit exceptions require reason and user confirmation; report exceptions rather than unqualified compliance.
- Check body-style coverage so untagged prose cannot evade typography checks. Headings, TOC, captions and source tables have separate roles; no whole-document cleanup.
- Verify appendix 1/2 source columns, appendix 3 merged stage header and unit人, appendix 4 readable schedule, appendix 5 layout and explanations, appendix 6 land area/location/time. Equipment/land facts remain evidence-backed or controlled prerequisites.
- For technical dark bids inspect in-chapter form placement and forbidden identity in body, images, headers, properties, comments and revisions. Do not assume every tender is dark bid.
- Latest profile: paired validator with `--structure-policy expandable --template ...`, profile auditor and protected-XML comparison. Legacy: fixed paired counts plus fidelity auditor. Record allowed changes; no blanket exemptions.

## Source and identity

- source files remain byte-identical;
- project name matches the user/source authority everywhere;
- stated duration matches the user instruction;
- stale project names, old durations, old paths, and obsolete standards are absent;
- output filenames and folder use the current project name.

## Word/WPS package

- DOCX is a valid ZIP package and opens with `python-docx`;
- WPS/Word opens, updates fields/TOC, saves, and exports PDF without repair prompts;
- counts remain stable for the legacy fixed template; latest-profile clones/row changes match the approved edit manifest;
- formatting/property drift outside approved slots is zero or explicitly justified;
- Chapter 1 covers every material BOQ package and the figure count/provenance matches the plan;
- no old project content remains in body, tables, headers/footers, relationships, comments, or core properties;
- no unintended blank/sparse, clipped, overlapped, or orphan-heading pages exist in the full render (blank reference rows/pages do not require empty project pages);
- TOC entries and page numbers are present after field update.

## Excel package

- valid XLSX ZIP, opens with `openpyxl` and WPS/Excel;
- exactly five required visible sheets in the required order;
- project name and duration are correct;
- Gantt rightmost day equals `n` and the completion milestone is on day `n`;
- network IDs/dependencies and resource phases reconcile with Gantt;
- BOQ summary item count equals the inspected source count;
- no formula-error value or missing intended formula result;
- every sheet has usable print area, scale, repeated titles, and a readable render.

## Manual visual sampling

Inspect at least:

- cover and approval/title pages;
- complete TOC;
- first and last Chapter 1 pages;
- every figure page;
- widest/highest-risk tables;
- chapter and section transitions;
- final Word page;
- all five Excel sheets, especially the rightmost schedule day.

## Final folder hygiene

The user-facing folder should normally contain only the final `.docx`, optional `.wps`, and `.xlsx`. Keep PDF renders, image assets, source extracts, builders, contract JSON, and audit reports in the project working/QA folder.

Do not declare delivery complete when any mandatory check fails. Report the exact failed check and the safest next action.
