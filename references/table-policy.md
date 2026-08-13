# Table Admission Policy

## Default Position

Prefer numbered prose. A construction scheme is an engineering explanation, not a spreadsheet transferred into Word. A table must improve comparison, lookup, responsibility assignment, or schedule reading. Visual neatness alone is not a reason to create one.

For a normal full scheme, target 6-10 body tables beyond fixed cover/approval furniture. This is a ceiling target, not a quota. A simpler project may need fewer.

## Admission Test

Create a table only when all answers are yes:

1. Does the content contain at least three repeated records?
2. Do those records share the same fields or measurement dimensions?
3. Will the reviewer compare rows or look up a value, owner, date, risk, or status?
4. Can each cell remain concise without hiding reasoning or a work sequence?
5. Is this information not already explained adequately in adjacent prose?

If any answer is no, use a heading with paragraphs, numbered steps, or a short list.

## Normally Retained Tables

- template approval and signature table;
- one compact organization/responsibility table when supplied names or roles require lookup;
- one major-quantity summary only when it helps define a broad BOQ scope;
- one consolidated hazard-identification matrix;
- one manpower plan;
- one combined machinery, tools, and test-instrument plan where practical;
- one inspection/test and hold-point plan;
- one emergency command/contact or emergency-material table when verified information exists;
- one landscape Gantt or horizontal-bar schedule;
- a review-comment response matrix only when the task specifically requires it.

Do not add all items automatically. Keep only those needed by the project.

## Content That Must Stay Prose

- compilation purpose and project interpretation;
- project characteristics, difficulties, constraints, and overall deployment;
- each construction method and operating sequence;
- technical parameters with their source and application logic;
- quality assurance, defect handling, and acceptance reasoning;
- major-risk controls beyond the concise hazard matrix;
- standards excerpts and project-specific interpretation;
- civilized construction, environment, occupational health, and seasonal measures;
- emergency procedures and recovery conditions;
- cross-discipline and owner/production/supervision coordination;
- finished-product protection, commissioning, handover, and warranty commitments.

## Anti-Patterns

- one two-column table per subsection;
- tables whose left column is only `项目/内容/要求` and right cells contain paragraphs;
- repeating the same prose in a table and adjacent text;
- splitting one resource, hazard, or inspection dataset into many small tables;
- using a table as a colored box, visual divider, or page-filling device;
- shrinking table type to fit content that should have been prose.

## Automated Gate

The depth audit counts all DOCX tables, including fixed front matter, so its density limits allow reasonable template furniture. For `standard` and `complex` profiles:

- explanatory text share: at least 82%;
- tables: at most 4 per 10,000 body characters;
- table rows: at most 35 per 10,000 body characters;
- prose-heavy tables: at most two for `standard` and three for `complex`.

A prose-heavy table has multiple long cells or an excessive average cell length. Passing these numeric limits does not excuse an unnecessary table; visual review must still apply the admission test.
