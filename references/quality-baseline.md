# Full-Depth Quality Baseline

## Reference Standard

The accepted cold-rolling 101 scheme is bundled as `assets/冷轧厂101完整度质量基准.docx`. Use it as the depth reference, not as a body-text source to copy blindly. Its measured characteristics were approximately 52,820 body characters, 1,631 nonempty paragraphs, 308 heading candidates, 15 major chapters, and broad technical plus management coverage. Later short schemes with only about 9,000-11,000 body characters were visually polished but materially below this standard.

Use the reference to preserve these qualities:

- complete construction and management lifecycle;
- project-specific work-package decomposition;
- detailed method, inspection, safety, record, and handover logic;
- restrained use of tables and standards excerpts;
- enough management content for owner, supervision, production, safety, quality, and completion review.

Do not force every small project to match the reference word for word. Select one profile and pass its deterministic gate.

## Profiles

| Profile | Use | Body chars | Explanatory chars | Explanatory paragraphs | Headings | Chapters | Management topics | Project terms | Method terms |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|
| `small` | Explicitly requested concise, single-discipline, low-risk work only | 22,000 | 18,000 | 180 | 50 | 10 | 10 | 4 | 5 |
| `standard` | Default industrial construction scheme | 35,000 | 30,000 | 280 | 85 | 14 | 12 | 6 | 8 |
| `complex` | Multi-discipline, shutdown, hazardous media, lifting, confined space, pressure, or major production interface | 48,000 | 40,000 | 360 | 120 | 15 | 14 | 8 | 12 |

The script also enforces:

- explanatory text is at least 82% of explanatory plus table text for `standard` and `complex`;
- total tables do not exceed 4 per 10,000 body characters and table rows do not exceed 35 per 10,000 body characters for `standard` and `complex`;
- prose-heavy tables remain exceptional: no more than two in `standard` and three in `complex`, normally limited to consolidated hazards or inspections;
- at least 1,800 narrative characters per embedded media item for `standard` and `complex`, allowing one cover/logo asset;
- at least 4,000 narrative characters per standards-excerpt caption;
- project terms occur repeatedly and are distributed through multiple major chapters;
- a meaningful share of method terms is surrounded by sequence, quality, safety, records, and acceptance language;
- legacy project terms are absent;
- duplicate long paragraphs remain below the profile limit.

## Profile Selection

- Choose `complex` when any two complexity indicators apply, or when one high-consequence indicator such as hazardous-media tie-in, major lifting, confined-space rescue exposure, or pressure-system modification dominates the scope.
- Choose `standard` by default, including ordinary multi-trade building, electrical, piping, landscape, and repair projects.
- Choose `small` only after recording the user's explicit concise request and confirming that no complex indicator applies. Short duration or small contract value alone does not justify `small`.

## No Threshold Gaming

Never lower thresholds, repeat generic prose, split one paragraph into many, add empty headings, or insert decorative images to obtain a pass. If a legitimate project cannot meet a default metric, document the engineering reason and obtain an explicit user-approved exception; keep all other gates active.
