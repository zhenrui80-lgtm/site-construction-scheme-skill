# End-to-end workflow

## 1. Intake and isolation

- Create one working directory per project, for example `_work_<slug>_<date>`. Never modify a prior builder in place.
- Copy sources into `intake/`; record size, modified time, and SHA-256. Preserve original formats.
- Inventory Word, Excel, PDF, RAR/ZIP, CAD, photos, and existing schemes. Extract archives with an RAR5-capable 7-Zip when required.
- Prefer ASCII temporary paths if a library fails on Chinese paths; keep the final output path in Chinese.

## 2. Evidence extraction

- Inspect every workbook sheet, including hidden rows/columns, merged cells, formulas, units, quantities, notes, and subtotal rows.
- Read tender documents and task books for project name, scope boundary, duration, qualification/response clauses, milestones, shutdown windows, owner-supplied items, acceptance records, and prohibited assumptions.
- Read current format requirements for chapter coverage, attached forms, page-count scope, typography and technical dark-bid conditions. Use the newest applicable user-authorized format, not an old project's visual habits.
- Inspect drawings and photos visually. When the user says “从图纸里面扣”, use cropped source-drawing extracts with captions instead of generic illustrations.
- Build a source register with `source`, `fact`, `location`, `confidence`, and `impact`. Mark conflicts for review.

## 3. Classification

Choose one mode before writing:

- `construction`: installation, repair, demolition, civil, piping, steelwork, electrical, instrumentation, commissioning, or physical handover dominates.
- `service`: inspection, operation, maintenance support, cleaning, transport, staffing, consulting, records, response/SLA, or recurring service delivery dominates.
- `mixed`: both are material. Separate the service organization from the physical work packages and keep one interface matrix.

## 4. Project contract

Record at minimum:

- exact project name and accepted filename-safe variant;
- exact user-stated duration and basis (`D1…Dn` unless fixed dates are explicitly given);
- BOQ item count and normalized work-package register;
- source files and hashes;
- template path/hash and permitted edit slots;
- format profile (`latest-20261023`, explicit `legacy-transport`, or project-specific), body styles, protected appendix schemas, allowed clones and approved deviations;
- page limit/counting scope, rendered count, and any user-confirmed exception; dark-bid status and identity-scrub scope;
- 15-topic/six-appendix mapping for latest construction mode, or tender-specific service mapping;
- output DOCX/WPS/XLSX/PDF paths;
- stale names, stale durations, and obsolete standards to reject;
- work fronts, dependencies, milestones, resources, hazards, interfaces, and evidence gaps;
- expected figure list and figure provenance;
- validation thresholds.

## 5. Authoring

- Write Chapter 1 first from the BOQ register, then reconcile remaining chapters with it.
- Use a fresh builder. Shared utilities may be reused, but project constants and content must live in the new project contract/builder.
- Replace text in existing runs; preserve the run containing a TOC tab or leader where present.
- Keep numeric TOC field results in their original numeric run, not a compressed spacing/tab run. Fill merged table cells once per unique XML cell; never overwrite one vertically merged cell through each logical row.
- Under the latest profile clone only documented styles/data rows. Keep Chapter 1 detailed and remove repetition elsewhere to stay within the 200-page target; do not change body typography to force a page count.
- Insert figures only in approved body paragraphs. Give every figure a source/provenance note and a clear caption.
- Keep unknown values as controlled prerequisites or hold points; do not invent them.

## 6. Schedule and workbook

- Use the same work-package IDs, durations, dependencies, and milestones in Word and Excel.
- Make the critical path credible for the stated scope and duration; parallelize only when separate work fronts/resources make it feasible.
- Use `D1…Dn` when no start date is supplied. Do not invent calendar dates.
- End with a completion/acceptance milestone on exactly day `n`.

## 7. QA and delivery

- Run structural, content, schedule, residue, and package-integrity checks.
- Run the selected-profile format audit after the final field update; cached Word page counts are not evidence. Latest profile checks body宋体12pt/1.5, 15 topics and six appendices; legacy fidelity must not exempt old figure slots for another hash.
- Open and save through WPS/Word; update TOC fields and export PDF.
- Render every PDF page and every workbook sheet. Inspect cover, TOC, Chapter 1, figure pages, tables, section transitions, and the last schedule column.
- Deliver only after all required checks pass. Put only the final DOCX/WPS/XLSX in the user-facing project folder unless the user asks for QA artifacts.
