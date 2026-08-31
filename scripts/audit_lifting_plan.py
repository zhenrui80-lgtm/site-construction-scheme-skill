#!/usr/bin/env python3
"""Audit the data, calculations, evidence, and Word equations of a lifting plan."""

from __future__ import annotations

import argparse
import json
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    from .project_contract import load_contract, resolve_path
except ImportError:  # Script execution from the skill root.
    from project_contract import load_contract, resolve_path


CHECKS = (
    "design_lifting_load",
    "crane_capacity",
    "rigging",
    "lifting_points",
    "spreader_beam",
    "outrigger_reaction",
    "ground_bearing",
    "station_setout",
    "clearance_and_hook_height",
    "trial_lift",
    "wind_control",
)
FIGURES = ("lifting_point_detail", "station_layout", "outrigger_bearing", "rigging_configuration")
STATUSES = {"verified", "derived", "site_confirmation", "not_applicable"}
CHECK_VALUE_FIELDS = {
    "design_lifting_load": ("governing_component_weight_t", "additional_lifting_weight_t", "design_lifting_weight_t"),
    "crane_capacity": ("design_lifting_weight_t", "rated_capacity_t", "capacity_utilization"),
    "rigging": ("maximum_branch_force_kn", "selected_capacity_kn"),
    "lifting_points": ("maximum_point_force_kn", "selected_capacity_kn"),
    "spreader_beam": ("design_force_kn", "utilization"),
    "outrigger_reaction": ("maximum_reaction_kn",),
    "ground_bearing": ("maximum_reaction_kn", "effective_area_m2", "calculated_pressure_kpa", "allowable_bearing_kpa"),
    "station_setout": ("working_radius_m", "radius_tolerance_m"),
    "clearance_and_hook_height": ("required_hook_height_m", "available_hook_height_m"),
    "trial_lift": ("trial_lift_height_mm",),
    "wind_control": ("stop_wind_speed_m_s",),
}
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def positive_number(value):
    try:
        return float(value) > 0
    except (TypeError, ValueError):
        return False


def close_enough(left, right, relative_tolerance=0.03):
    return abs(left - right) <= max(abs(left), abs(right), 1.0) * relative_tolerance


def resolve_evidence_path(base, value):
    path = Path(str(value)).expanduser()
    return path.resolve() if path.is_absolute() else (base / path).resolve()


def audit_source(source, label, base, failures, execution_ready=True):
    if not isinstance(source, dict):
        failures.append(f"{label} source must be an object")
        return
    if execution_ready and source.get("status") != "verified":
        failures.append(f"{label} source status must be verified")
    filename = str(source.get("file", "")).strip()
    if not filename:
        failures.append(f"{label} source.file is required")
    elif not resolve_evidence_path(base, filename).exists():
        failures.append(f"{label} source file does not exist: {filename}")
    if not any(str(source.get(key, "")).strip() for key in ("page", "location", "drawing_no", "record_no")):
        failures.append(f"{label} source requires page, location, drawing_no, or record_no")


def audit_docx_equations(docx, minimum):
    result = {"docx": str(docx), "minimum": minimum, "equations": 0, "times_new_roman_equations": 0}
    failures = []
    if not docx.exists():
        return result, [f"lifting plan DOCX does not exist: {docx}"]
    try:
        with zipfile.ZipFile(docx) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
    except (zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
        return result, [f"cannot inspect lifting plan equations: {exc}"]
    equations = root.findall(f".//{{{MATH_NS}}}oMath")
    result["equations"] = len(equations)
    for equation in equations:
        fonts = equation.findall(f".//{{{WORD_NS}}}rFonts")
        if fonts and all(
            node.get(f"{{{WORD_NS}}}{slot}") == "Times New Roman"
            for node in fonts
            for slot in ("ascii", "hAnsi", "eastAsia", "cs")
        ):
            result["times_new_roman_equations"] += 1
    if len(equations) < minimum:
        failures.append(f"editable OMML equations {len(equations)} < required {minimum}")
    if equations and result["times_new_roman_equations"] != len(equations):
        failures.append("every OMML equation must use Times New Roman in all Word font slots")
    return result, failures


def audit_lifting_plan(contract_path, contract, docx_path=None):
    failures = []
    special = contract.get("special_plan", {})
    if special.get("type") != "lifting":
        return {"applicable": False, "passed": True, "failures": []}
    execution_ready = special.get("delivery_mode", "execution_ready") == "execution_ready"
    if contract.get("document", {}).get("profile") != "complex":
        failures.append("lifting special plan requires document.profile=complex")
    register_value = special.get("data_register")
    if not register_value:
        failures.append("special_plan.data_register is required")
        return {"applicable": True, "passed": False, "failures": failures}
    register_path = resolve_path(contract_path, register_value)
    if not register_path.exists():
        failures.append(f"lifting data register does not exist: {register_path}")
        return {"applicable": True, "data_register": str(register_path), "passed": False, "failures": failures}
    try:
        register = json.loads(register_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        failures.append(f"cannot read lifting data register: {exc}")
        return {"applicable": True, "data_register": str(register_path), "passed": False, "failures": failures}
    if register.get("schema_version") != 1:
        failures.append("lifting data register schema_version must be 1")
    if register.get("delivery_mode") != special.get("delivery_mode", "execution_ready"):
        failures.append("lifting data register delivery_mode must match project contract")
    if execution_ready and register.get("unresolved_items"):
        failures.append("execution-ready lifting plan cannot contain unresolved_items")

    components = register.get("components", [])
    if not components:
        failures.append("at least one lifted component is required")
    for index, component in enumerate(components, start=1):
        label = f"components[{index}]"
        if not str(component.get("id", "")).strip():
            failures.append(f"{label}.id is required")
        if not positive_number(component.get("actual_weight_t")):
            failures.append(f"{label}.actual_weight_t must be positive")
        audit_source(component.get("weight_source"), f"{label}.weight", register_path.parent, failures, execution_ready)
        cog = component.get("center_of_gravity")
        if not isinstance(cog, dict) or cog.get("status") not in STATUSES - {"not_applicable"}:
            failures.append(f"{label}.center_of_gravity requires a valid status")
        else:
            if execution_ready and cog.get("status") == "site_confirmation":
                failures.append(f"{label}.center_of_gravity remains site_confirmation")
            if not str(cog.get("method", "")).strip():
                failures.append(f"{label}.center_of_gravity.method is required")
            audit_source(cog.get("source"), f"{label}.center_of_gravity", register_path.parent, failures, execution_ready)

    configurations = register.get("crane_configurations", [])
    if not configurations:
        failures.append("at least one exact crane configuration is required")
    for index, config in enumerate(configurations, start=1):
        label = f"crane_configurations[{index}]"
        for field in ("id", "manufacturer", "model", "outrigger_configuration", "boom_configuration"):
            if not str(config.get(field, "")).strip():
                failures.append(f"{label}.{field} is required")
        for field in (
            "counterweight_t",
            "working_radius_m",
            "boom_length_m",
            "rated_capacity_t",
            "design_lifting_weight_t",
            "capacity_utilization",
        ):
            if not positive_number(config.get(field)):
                failures.append(f"{label}.{field} must be positive")
        if not config.get("governing_components"):
            failures.append(f"{label}.governing_components is required")
        if all(positive_number(config.get(field)) for field in ("rated_capacity_t", "design_lifting_weight_t", "capacity_utilization")):
            utilization = float(config["design_lifting_weight_t"]) / float(config["rated_capacity_t"])
            if not close_enough(utilization, float(config["capacity_utilization"])):
                failures.append(f"{label}.capacity_utilization does not equal design weight divided by rated capacity")
            if float(config["capacity_utilization"]) > 1:
                failures.append(f"{label}.capacity_utilization exceeds 1.0")
        audit_source(config.get("load_chart_source"), f"{label}.load_chart", register_path.parent, failures, execution_ready)

    checks = register.get("engineering_checks", {})
    for name in CHECKS:
        check = checks.get(name)
        label = f"engineering_checks.{name}"
        if not isinstance(check, dict) or check.get("status") not in STATUSES:
            failures.append(f"{label} requires a valid status")
            continue
        status = check["status"]
        if execution_ready and status == "site_confirmation":
            failures.append(f"{label} remains site_confirmation")
        if status == "not_applicable":
            if not str(check.get("reason", "")).strip():
                failures.append(f"{label}.reason is required when not_applicable")
            continue
        if not str(check.get("calculation_ref") or check.get("record_ref") or "").strip():
            failures.append(f"{label} requires calculation_ref or record_ref")
        values = check.get("calculation_values")
        if not isinstance(values, dict):
            failures.append(f"{label}.calculation_values is required")
        else:
            for field in CHECK_VALUE_FIELDS[name]:
                if not positive_number(values.get(field)):
                    failures.append(f"{label}.calculation_values.{field} must be positive")
        sources = check.get("source_refs", [])
        if not sources:
            failures.append(f"{label}.source_refs is required")
        for source_index, source in enumerate(sources, start=1):
            audit_source(source, f"{label}.source_refs[{source_index}]", register_path.parent, failures, execution_ready)

    design_values = checks.get("design_lifting_load", {}).get("calculation_values", {})
    if all(positive_number(design_values.get(key)) for key in CHECK_VALUE_FIELDS["design_lifting_load"]):
        if design_values["design_lifting_weight_t"] < (
            design_values["governing_component_weight_t"] + design_values["additional_lifting_weight_t"]
        ):
            failures.append("design lifting weight is smaller than component plus additional lifting weight")
    crane_values = checks.get("crane_capacity", {}).get("calculation_values", {})
    if all(positive_number(crane_values.get(key)) for key in CHECK_VALUE_FIELDS["crane_capacity"]):
        utilization = crane_values["design_lifting_weight_t"] / crane_values["rated_capacity_t"]
        if not close_enough(utilization, crane_values["capacity_utilization"]):
            failures.append("crane capacity utilization does not equal design weight divided by rated capacity")
        if crane_values["capacity_utilization"] > 1:
            failures.append("crane capacity utilization exceeds 1.0")
    for name in ("rigging", "lifting_points"):
        values = checks.get(name, {}).get("calculation_values", {})
        force_key = "maximum_branch_force_kn" if name == "rigging" else "maximum_point_force_kn"
        if positive_number(values.get(force_key)) and positive_number(values.get("selected_capacity_kn")):
            if values["selected_capacity_kn"] < values[force_key]:
                failures.append(f"{name} selected capacity is smaller than calculated force")
    spreader = checks.get("spreader_beam", {})
    spreader_values = spreader.get("calculation_values", {})
    if spreader.get("status") != "not_applicable" and positive_number(spreader_values.get("utilization")):
        if spreader_values["utilization"] > 1:
            failures.append("spreader beam utilization exceeds 1.0")
    ground_values = checks.get("ground_bearing", {}).get("calculation_values", {})
    if all(positive_number(ground_values.get(key)) for key in CHECK_VALUE_FIELDS["ground_bearing"]):
        pressure = ground_values["maximum_reaction_kn"] / ground_values["effective_area_m2"]
        if not close_enough(pressure, ground_values["calculated_pressure_kpa"]):
            failures.append("ground pressure does not equal maximum reaction divided by effective area")
        if ground_values["calculated_pressure_kpa"] > ground_values["allowable_bearing_kpa"]:
            failures.append("calculated ground pressure exceeds allowable bearing pressure")
    clearance_values = checks.get("clearance_and_hook_height", {}).get("calculation_values", {})
    if positive_number(clearance_values.get("required_hook_height_m")) and positive_number(
        clearance_values.get("available_hook_height_m")
    ):
        if clearance_values["available_hook_height_m"] < clearance_values["required_hook_height_m"]:
            failures.append("available hook height is smaller than required hook height")

    figures = register.get("figures", {})
    required_figures = list(FIGURES)
    if checks.get("spreader_beam", {}).get("status") != "not_applicable":
        required_figures.append("spreader_beam_detail")
    for name in required_figures:
        figure = figures.get(name)
        label = f"figures.{name}"
        if not isinstance(figure, dict):
            failures.append(f"{label} is required")
            continue
        filename = str(figure.get("file", "")).strip()
        if not filename:
            failures.append(f"{label}.file is required")
        elif not resolve_evidence_path(register_path.parent, filename).exists():
            failures.append(f"{label} file does not exist: {filename}")
        if execution_ready and figure.get("status") not in {"verified", "derived"}:
            failures.append(f"{label}.status must be verified or derived")

    minimum = int(special.get("minimum_omml_equations", 8))
    docx = Path(docx_path).resolve() if docx_path else resolve_path(contract_path, contract["document"]["output_docx"])
    equation_audit, equation_failures = audit_docx_equations(docx, minimum)
    failures.extend(equation_failures)
    return {
        "applicable": True,
        "delivery_mode": special.get("delivery_mode", "execution_ready"),
        "data_register": str(register_path),
        "components": len(components),
        "crane_configurations": len(configurations),
        "engineering_checks": len(checks),
        "equations": equation_audit,
        "failures": failures,
        "passed": not failures,
    }


def main():
    parser = argparse.ArgumentParser(description="Audit a lifting special plan and its evidence register.")
    parser.add_argument("contract", type=Path)
    parser.add_argument("--docx", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    contract_path, contract = load_contract(args.contract)
    result = audit_lifting_plan(contract_path, contract, args.docx)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    raise SystemExit(0 if result["passed"] else 1)


if __name__ == "__main__":
    main()
