# Excel schedule and response workbook

Create one `.xlsx` with exactly five visible worksheets in this order:

1. `横道进度计划`
2. `网络计划逻辑`
3. `资源计划`
4. `评标响应索引`
5. `清单工程量摘要`

Do not add scratch or chart-data sheets to the final workbook.

## 横道进度计划

- Use a fixed left register: ID, activity, work package, duration, start day, finish day, predecessor, responsibility, and remarks.
- Use day-index columns `D1…Dn` when the user supplies only a duration. Do not invent dates.
- Include mobilization/preparation, every material work package, testing/commissioning, punch-list/rectification, document handover, and a zero/one-day final milestone as appropriate.
- Place the final milestone on day `n`. The rightmost plan header must also represent day `n`.
- Use print titles, frozen panes, legible colors, landscape layout, fit-to-width, and explicit print area. Do not hide the final day.

## 网络计划逻辑

- Use the same activity IDs as the Gantt sheet.
- Record predecessor, successor, relationship, lag, critical/non-critical flag, interface/hold point, and logic rationale.
- Show a coherent path from award/preparation through acceptance. Parallel activities require distinct work fronts/resources or an explicit interface plan.

## 资源计划

- Plan labor, management, main plant/tools, temporary facilities, testing instruments, and peak demand by phase/day range.
- Reconcile key resources with the methods and schedule. Do not schedule the same unique resource on incompatible concurrent fronts.
- Mark owner-provided resources or values to be confirmed.

## 评标响应索引

- Map each known tender/task-book requirement to the relevant Word chapter/page placeholder and supporting evidence.
- Include requirement source and location, response summary, document location, evidence/record, responsible role, and compliance status.
- If formal evaluation criteria were not supplied, state that the index is a traceability aid and must be updated after final tender-document review. Do not invent scoring clauses.

## 清单工程量摘要

- Preserve every BOQ item number, description, unit, quantity, key note, normalized work package, and Word method reference.
- Reconcile source row count and material quantities. Keep formulas auditable and avoid hard-coded summary totals where formulas are appropriate.

## Workbook QA

- Open with formulas and with cached values; flag `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#N/A`, and formulas lacking intended cached results.
- Confirm the exact five sheet names/order, no hidden sheets, no stale project names, the correct project name/duration, and the last day/final milestone equal `n`.
- Render every sheet and visually inspect page breaks, repeated titles, clipped columns, legends, the last day, and print scale.
