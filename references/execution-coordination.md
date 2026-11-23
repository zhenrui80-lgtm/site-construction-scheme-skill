# Execution coordination for construction schemes

This changes scheduling and evidence reuse only. All requirements in SKILL.md and its existing references remain binding. Complete the same full scheme, calculations, evidence, and release gates whether execution is serial or delegated.

## One intake, reusable evidence

- Complete the existing intake, template inspection, BOQ extraction, and project-contract steps. Record source identity/hash and exact sheet/row, drawing/view, PDF page, or photograph location with each extracted fact. Reuse `boq-analysis.json`, the coverage matrix, and evidence files; do not ask every author to re-extract the complete source set.
- Inspect every supplied source as required. A shared inventory is a coverage checklist, not a substitute for reading drawings, legends, details, notes, or uncertain entries. Reviewers verify cited raw evidence independently where conclusions depend on it.
- Read the supporting guidance at the stages already specified in SKILL.md. Provide a worker only its relevant references and source slices, plus the selected template/depth contract and complete interface constraints. Lifting workers must read the lifting reference; every technical author must apply the content checklist and table policy. The coordinator still reads and enforces all applicable guidance.
- Reuse a PDF extraction or page render only when source hash, page/crop, and rendering settings match. Existing archived standards may avoid duplicate PDF downloads, but never replace current-status verification or project-specific applicability review. Record verification date and source once per standard in this task.

## Bounded parallel work

After intake and the project coverage contract are established, assign substantial independent tasks when delegation is available:

| Work package | Required input | Return for coordinator review |
|---|---|---|
| Discipline method drafts | Assigned BOQ rows, drawings, eight-element checklist, interfaces | Complete numbered method prose, covered row IDs, source locations, unresolved facts, handover/hold points |
| Standards research | Assigned disciplines/hazards and current evidence register | Official status evidence, exact clause/page, provenance, applicability, downloaded file if available |
| Drawing/figure evidence | Assigned drawing/view IDs and figure purpose | Source crops or figure specification, verified dimensions, caption/provenance, unresolved geometry |
| Specialist calculation checks | Verified component/configuration input set and calculation references | Reproducible calculation or independent check, units, source references, acceptance criterion, discrepancies |
| Read-only final review | Frozen final artifacts, source register and gate reports | Assigned coverage/technical/visual findings with precise locations, including an explicit no-finding result |

Each task owns separate draft/evidence/report paths. Specify its scope, input revision, dependencies, and completion criteria; require facts and source locations rather than only a summary. Do not dispatch empty roles or one agent per short paragraph. The coordinator integrates cross-discipline permits, isolation, resource conflicts, acceptance responsibility and schedule before writing the final document.

## Dependencies and single ownership

- Only the coordinator updates the canonical project contract, coverage matrix, builder, final DOCX, shared QA summaries, and standard-library index. Workers submit proposed entries; they do not edit shared outputs or template files.
- For a newly downloaded standard, coordinate immediate `archive_standard.py` execution serially by the library owner before any later work on that PDF. A worker must not batch unarchived PDFs for later use or race another writer to `standards/index.json`. Subsequent extraction/rendering uses the archived file.
- Lifting remains a dependency chain: source weight/COG, geometry and ground evidence → completed register → governing load and exact crane configuration → dependent rigging/points/reactions/ground/clearance checks → checked figures and field controls. Parallelize only checks with all inputs fixed and no dependency on each other's unfinished result. A calculation cannot be its own independent check.
- Missing execution inputs, user-approved exceptions and template revisions retain the original stop/confirmation conditions. Other independent evidence gathering can proceed; dependent drafting or release cannot cross that gate. Do not invent values to keep a worker busy.
- One writer assembles the DOCX using the existing template/helper rules, then runs `validate_scheme.py` intact. Do not parallelize that command with document editing or Word field update/export, and do not replace it with selected component audits or `--skip-word-export`.

## Review and revisions

After final generation/export, assign disjoint read-only review ranges if useful. Account for every required contact sheet and every required full-size critical page, including all lifting calculation and figure checks. Parallel review does not reduce visual coverage or engineering depth.

Merge findings before editing. Changed inputs invalidate dependent prose, calculations, figures, schedules and reports; refresh those branches. Regenerate changed artifacts and run the complete final gate and required visual review on the resulting final revision. Old page numbers, old renders and a worker's success report cannot certify a changed DOCX. Preserve the original handoff fields and disclose unresolved items.
