# Lifting Special-Plan Workflow

Use this workflow for a standalone 专项吊装方案 and for any construction scheme
whose lifting chapter selects a crane, defines lifting points or rigging, uses a
spreader beam, checks outrigger reactions, or fixes a crane standing position.
Always use the `complex` profile.

## 1. Initialize The Mode

```powershell
python scripts/init_project.py "<workdir>" --project "<project>" --template-style 冷轧厂 --duration-days 16 --special-plan lifting
```

The command creates `project-contract.json` and an intentionally incomplete
`lifting-data-register.json`. The incomplete register is a gate, not boilerplate:
the plan cannot pass final validation until every execution input is closed.

Use `special_plan.delivery_mode=execution_ready` only for a scheme that can be
issued for field execution. `preliminary_review` may be used while gathering
inputs, but label the document and handoff clearly and do not describe it as an
approved or executable final plan.

## 2. Establish A Source Hierarchy

Use sources in this order and retain the files in the project evidence folder:

1. approved design and fabrication drawings, BOMs, shipping lists, weighbridge or load-cell records;
2. written design/manufacturer responses and approved changes;
3. site survey, control-point, elevation, obstruction, underground-utility, and standing-area records;
4. geotechnical report, plate-load test, concrete assessment, or owner-approved ground capacity record;
5. exact crane manufacturer's manual, configuration sheet, load chart, and outrigger data;
6. current official national, industry, ministry, and special-equipment standards.

Web research may locate current official standards and the exact manufacturer's
documents. It must not supply a project component's weight, center of gravity,
lifting-point geometry, ground capacity, underground condition, or standing
coordinate. If those project facts are missing, stop the execution-ready issue,
record them in `unresolved_items`, and obtain written or measured evidence.

Classify each datum as:

- `verified`: read directly from a controlled source or measurement record;
- `derived`: calculated from verified inputs, with a calculation reference;
- `site_confirmation`: not yet closed and therefore not execution-ready;
- `not_applicable`: genuinely unused, with a technical reason.

Never change `site_confirmation` to `verified` merely to pass the audit.

## 3. Complete The Lifting Data Register

### Components

Create one entry per lifted item or fabrication segment. Record:

- unique component ID and drawing/BOM identity;
- actual controlled weight in tonnes, including preassembled attachments;
- weight source file and drawing/page/record location;
- center-of-gravity status, method, coordinates or described datum, and source;
- overall dimensions, lifting orientation, installation elevation, and governing interfaces;
- temporary attachments that are included in the design lifting weight.

An area total from the BOQ is not a component weight. Proportional allocation by
length or area may be used only for a preliminary sensitivity study and remains
`site_confirmation`; it cannot support an execution-ready lift.

### Exact Crane Configurations

Create one entry for every governing configuration. Record the exact:

- manufacturer, model, rated class, equipment number when known;
- counterweight, superlift, jib, main-boom, hook block and reeving configuration;
- outrigger extension and operating sector;
- measured working radius and boom length;
- design lifting weight, original-chart rated capacity, and utilization ratio;
- governing components and the original manufacturer file plus page or table location.

Do not write only “240 t class crane” or substitute a similar model's chart.
Model, configuration, unit, sector, radius interval, and chart note must all match.
Do not interpolate between chart cells unless the manufacturer explicitly permits it.

### Engineering Checks

Complete every check in the register:

1. `design_lifting_load`: actual component, hook block, spreader, rigging, attachments, dynamic and imbalance effects;
2. `crane_capacity`: exact load-chart selection, utilization, boom and radius control;
3. `rigging`: sling geometry, branch forces, WLL, angle, bend efficiency, shackles and hook;
4. `lifting_points`: force distribution from actual center of gravity, lug plate, pin, weld and parent-member checks;
5. `spreader_beam`: bending, shear, axial force, stability, local plate, lug, pin and weld checks, or a reason for not using one;
6. `outrigger_reaction`: manufacturer's output or accepted engineering method for the worst slew direction, not total load divided by four;
7. `ground_bearing`: effective pad area, local plate/mat strength, allowable bearing, settlement, edge distance and buried facilities;
8. `station_setout`: control points, coordinates or offsets, outrigger corners, radius tolerance, travel/slew envelope and no-go boundary;
9. `clearance_and_hook_height`: component depth, rigging headroom, hook block, boom geometry, obstacle and landing allowance;
10. `trial_lift`: 100-200 mm trial-lift hold point, brake, balance, deformation, settlement, alarm and release criteria;
11. `wind_control`: component wind area, manufacturer and standard limits, field instrument, stop-work and securing method.

Every check needs verified source references plus a calculation or record reference.
Also fill its structured `calculation_values`; the audit checks basic arithmetic
consistency for crane utilization, selected rigging and lifting-point capacity,
spreader utilization, outrigger reaction/ground pressure, and hook-height margin.
These machine checks do not replace the signed calculation book or an independent
engineering review.
Use `not_applicable` only with a reason. An execution-ready register cannot
contain a core check with `site_confirmation` status or unresolved items.

## 4. Write A Complete Calculation Book

For each calculation, show these items in order:

1. calculation object and governing component/configuration;
2. input table with value, unit, status, source file and page/drawing/record;
3. symbols and units;
4. design combination and coefficient basis;
5. formula as an editable Word equation;
6. numeric substitution with consistent units;
7. result and rounding rule;
8. limit or acceptance criterion;
9. conclusion and affected construction control;
10. preparer and independent checker.

Typical relationships include the following, but coefficients and limits must
come from the applicable standard, manufacturer, design, or approved project rule:

- design lifting load: `Q_d = gamma_d × (G_c + G_a + G_r)`;
- crane utilization: `eta = Q_d / Q_r`;
- sling force from the actual branch geometry and unequal load distribution;
- lifting-point reactions from component equilibrium about the verified center of gravity;
- maximum bearing pressure: `p_max = R_max / A_eff` and `p_max <= f_a`;
- spreader-beam force, bending, shear, combined stress, stability, lug, pin and weld checks;
- hook-height and boom-clearance geometry at the measured radius.

Do not use a simplified formula when it conceals the governing eccentricity,
operating sector, wind action, load distribution, or local structural check.
Show intermediate results so another engineer can reproduce the conclusion.

## 5. Use Editable Word Equations

Import `add_omml_equation()` from `scripts/docx_builder.py` and use it for every
display formula. The helper creates Word OMML equation objects and sets all math
font slots to Times New Roman. It accepts simple subscript notation such as
`Q_{d}` and `R_{max}`.

```python
from scripts.docx_builder import add_omml_equation

add_omml_equation(
    doc,
    "eta = Q_{d} / Q_{r} = 39.62 / 49.53 = 0.800",
    "式中各输入值见本工况输入表，结论按项目批准控制值判定。",
)
```

Do not paste equation screenshots, use an image of a calculation sheet, or leave
garbled Unicode formula text. The final audit opens `word/document.xml`, counts
`m:oMath` objects, and verifies Times New Roman in every equation font slot.

## 6. Produce Engineering Figures

At minimum provide:

- lifting-point detail tied to component ID, center of gravity, dimensions, and force direction;
- crane standing/setout plan tied to survey control, radius, outrigger corners, obstacles and exclusion zone;
- outrigger reaction and pad/ground bearing diagram;
- rigging configuration with hook, spreader when used, sling branches, shackles and component points;
- spreader-beam detail and force path when a spreader is used;
- elevation/clearance profile for high or obstructed lifts.

Derive figures from controlled drawings, survey, manufacturer data, or checked
calculations. Generated graphics may clarify sequence or force paths but cannot
invent dimensions. Label an unscaled generated image “示意，尺寸及位置以审定图纸和现场放样为准”.

## 7. Build The Field-Control Chapters

The body must connect the calculations to field actions:

- work boundary, production interface, permits and lifting classification;
- organization, named roles only when supplied, command hierarchy and stop-work authority;
- component receipt, identity, weight/COG closure and lifting-point inspection;
- crane entry, certificates, assembly/configuration, setout and pre-use inspection;
- ground treatment, pads, outrigger setup and settlement monitoring;
- rigging inspection, trial assembly, attachment, trial lift and formal lift;
- slewing, tag lines, landing, temporary stability, connection and staged unloading;
- wind, rain, visibility, night-work and production-alarm limits;
- high work, hot work, electrical isolation, falling-object and crossing-operation controls;
- quality hold points, single-lift work card, monitoring and traceable records;
- lifting-specific hazard identification, emergency lowering, instability, collision, outrigger settlement, weather and rescue response;
- change control: any change to component, weight, COG, point, sling, crane, configuration, radius, standing area, ground, sequence or weather boundary invalidates the affected work card and calculation.

Use tables only for comparison records such as component cards, crane work cards,
hazard registers, hold points, and a schedule. Keep the method and management
logic mainly in prose.

## 8. Run The Delivery Gate

```powershell
python scripts/validate_scheme.py "<workdir>\project-contract.json"
```

For lifting mode this adds `qa/lifting-plan-audit.json` to the normal V2 checks.
Delivery is blocked if the register, source files, component weight/COG, exact
crane configuration, engineering checks, figures, minimum editable equations,
or Times New Roman equation fonts fail.

Before handoff, manually compare every number in the component cards, crane work
cards, load-chart excerpts, calculations, figures, method text, and inspection
forms. Search the entire document for stale crane models, obsolete capacities,
old component weights, copied project names, unsupported “estimated” values, and
conflicting units. Inspect calculation pages, figure pages, landscape pages, and
the approval/front-matter region at full rendered size.
