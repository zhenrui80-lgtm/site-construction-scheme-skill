# V2 Project Contract

Use one `project-contract.json` as the machine-readable source of truth for each
scheme. Keep project facts, template choice, schedule basis, coverage terms,
validation exceptions, and output paths in this file instead of scattering them
through a project-specific Python script.

## Initialize A Project

For fixed dates:

```powershell
python scripts/init_project.py "<workdir>" --project "<project>" --template-style 条材厂 --start-date 2026-08-14 --finish-date 2026-12-20 --profile standard
```

For a duration without confirmed calendar dates:

```powershell
python scripts/init_project.py "<workdir>" --project "<project>" --template-style 冷轧厂 --duration-days 20 --profile standard
```

The initializer creates `intake`, `assets`, `output`, and `qa`, copies the
selected template, and writes `project-contract.json`. Update the contract after
BOQ and drawing analysis and before document generation.

## Required Contract Decisions

- `project`: exact project name, contractor, plant/style, and depth profile.
- `schedule`: use `fixed_dates` when start and finish dates are supplied; use
  `day_index` only when the user gives duration alone. Fixed dates are inclusive.
- `document`: copied template path, project builder path, first body marker, and marker occurrence.
- `outputs`: final DOCX and PDF paths.
- `coverage`: project terms, method terms, forbidden legacy terms, disciplines,
  high-risk operations, and required content modules.
- `qa`: body font, render thresholds, approved front-matter text exceptions,
  and whether the final schedule page must be landscape.
- `standards`: versioned library root, required archived PDFs, verification
  switch, and optional strict metadata validation.

The coverage minimums are profile-dependent and enforced by
`scripts/project_contract.py`. The default `standard` profile requires at least
6 project terms and 8 method terms. Use `complex` for shutdown, lifting,
confined-space, pressure, hazardous-media, or broad multi-discipline work. Do
not weaken coverage merely to pass validation.

## Project Builder Boundary

Project-specific builders may import stable helpers from
`scripts/docx_builder.py` and contract functions from
`scripts/project_contract.py`. `scripts/audit_project_builder.py` enforces that
boundary during final validation. Builders must not import another project's
`build_scheme.py`, use `PREVIOUS_UTILS` or `BASE_BUILDER`, or load an absolute
path below a prior `work` directory. Copying a previous front matter is handled
through the declared template, not through Python inheritance.

## Single-Command Gate

After the DOCX has been assembled, run:

```powershell
python scripts/validate_scheme.py "<workdir>\project-contract.json"
```

The command updates Word fields, exports PDF, audits frozen front matter,
checks body-scoped font and required/forbidden terms, enforces depth and
schedule consistency, renders every PDF page, creates contact sheets, verifies
the standard library, and writes `qa/validation-summary.json`.

`--skip-word-export` exists for automated regression tests with an explicitly
declared existing PDF. Do not use it for delivery validation.

Automatic render checks detect blank pages, suspiciously sparse pages, likely
edge clipping, and missing landscape schedule pages. A person must still inspect
all contact sheets and the critical pages listed in `references/verification.md`.
Record any approved render exception narrowly in the contract; never reduce a
global threshold to conceal an incomplete document.
