import json
import tempfile
import unittest
import zipfile
from pathlib import Path

from scripts.audit_lifting_plan import CHECKS, audit_lifting_plan


DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body>{equations}</w:body></w:document>
"""


def equation(font="Times New Roman"):
    return (
        '<m:oMath><m:r><w:rPr><w:rFonts w:ascii="{0}" w:hAnsi="{0}" '
        'w:eastAsia="{0}" w:cs="{0}"/></w:rPr><m:t>Q_d=1</m:t></m:r></m:oMath>'
    ).format(font)


class LiftingPlanAuditTests(unittest.TestCase):
    def make_case(self, root):
        evidence = root / "evidence.pdf"
        evidence.write_bytes(b"evidence")
        for name in ("lifting-point.png", "station.png", "outrigger.png", "rigging.png"):
            (root / name).write_bytes(b"figure")
        docx = root / "plan.docx"
        with zipfile.ZipFile(docx, "w") as archive:
            archive.writestr("word/document.xml", DOCUMENT_XML.format(equations=equation() * 8))
        source = {"status": "verified", "file": evidence.name, "page": 1}
        checks = {}
        values = {
            "design_lifting_load": {
                "governing_component_weight_t": 32.02,
                "additional_lifting_weight_t": 4.0,
                "design_lifting_weight_t": 39.62,
            },
            "crane_capacity": {
                "design_lifting_weight_t": 39.62,
                "rated_capacity_t": 50.0,
                "capacity_utilization": 0.7924,
            },
            "rigging": {"maximum_branch_force_kn": 210.0, "selected_capacity_kn": 350.0},
            "lifting_points": {"maximum_point_force_kn": 130.0, "selected_capacity_kn": 220.0},
            "outrigger_reaction": {"maximum_reaction_kn": 2400.0},
            "ground_bearing": {
                "maximum_reaction_kn": 2400.0,
                "effective_area_m2": 12.0,
                "calculated_pressure_kpa": 200.0,
                "allowable_bearing_kpa": 220.0,
            },
            "station_setout": {"working_radius_m": 18.0, "radius_tolerance_m": 0.5},
            "clearance_and_hook_height": {"required_hook_height_m": 45.0, "available_hook_height_m": 48.0},
            "trial_lift": {"trial_lift_height_mm": 150.0},
            "wind_control": {"stop_wind_speed_m_s": 8.0},
        }
        for name in CHECKS:
            if name == "spreader_beam":
                checks[name] = {"status": "not_applicable", "reason": "本工况不使用平衡梁"}
            else:
                checks[name] = {
                    "status": "derived",
                    "calculation_ref": f"计算书-{name}",
                    "calculation_values": values[name],
                    "source_refs": [source],
                }
        register = {
            "schema_version": 1,
            "delivery_mode": "execution_ready",
            "components": [
                {
                    "id": "S1",
                    "actual_weight_t": 32.02,
                    "weight_source": source,
                    "center_of_gravity": {
                        "status": "derived",
                        "method": "按审定深化图建模计算",
                        "source": source,
                    },
                }
            ],
            "crane_configurations": [
                {
                    "id": "S1-C1",
                    "manufacturer": "Test Manufacturer",
                    "model": "TEST300",
                    "counterweight_t": 100,
                    "outrigger_configuration": "full extension",
                    "boom_configuration": "main boom",
                    "working_radius_m": 18,
                    "boom_length_m": 48,
                    "rated_capacity_t": 50,
                    "design_lifting_weight_t": 39,
                    "capacity_utilization": 0.78,
                    "governing_components": ["S1"],
                    "load_chart_source": source,
                }
            ],
            "engineering_checks": checks,
            "figures": {
                "lifting_point_detail": {"file": "lifting-point.png", "status": "derived"},
                "station_layout": {"file": "station.png", "status": "verified"},
                "outrigger_bearing": {"file": "outrigger.png", "status": "derived"},
                "rigging_configuration": {"file": "rigging.png", "status": "derived"},
            },
            "unresolved_items": [],
        }
        register_path = root / "lifting-data-register.json"
        register_path.write_text(json.dumps(register, ensure_ascii=False), encoding="utf-8")
        contract = {
            "document": {"profile": "complex", "output_docx": docx.name},
            "special_plan": {
                "type": "lifting",
                "delivery_mode": "execution_ready",
                "data_register": register_path.name,
                "minimum_omml_equations": 8,
            },
        }
        return contract, register, register_path, docx

    def test_execution_ready_register_passes(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            contract, _, _, docx = self.make_case(root)
            result = audit_lifting_plan(root / "project-contract.json", contract, docx)
        self.assertTrue(result["passed"], result["failures"])
        self.assertEqual(result["equations"]["equations"], 8)

    def test_unresolved_execution_item_is_blocked(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            contract, register, register_path, docx = self.make_case(root)
            register["unresolved_items"] = ["地基承载力未取得"]
            register_path.write_text(json.dumps(register, ensure_ascii=False), encoding="utf-8")
            result = audit_lifting_plan(root / "project-contract.json", contract, docx)
        self.assertFalse(result["passed"])
        self.assertTrue(any("unresolved_items" in failure for failure in result["failures"]))

    def test_plain_or_wrong_font_equations_are_blocked(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            contract, _, _, docx = self.make_case(root)
            with zipfile.ZipFile(docx, "w") as archive:
                archive.writestr("word/document.xml", DOCUMENT_XML.format(equations=equation("Cambria Math") * 8))
            result = audit_lifting_plan(root / "project-contract.json", contract, docx)
        self.assertFalse(result["passed"])
        self.assertTrue(any("Times New Roman" in failure for failure in result["failures"]))

    def test_inconsistent_ground_pressure_is_blocked(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            contract, register, register_path, docx = self.make_case(root)
            register["engineering_checks"]["ground_bearing"]["calculation_values"]["calculated_pressure_kpa"] = 150.0
            register_path.write_text(json.dumps(register, ensure_ascii=False), encoding="utf-8")
            result = audit_lifting_plan(root / "project-contract.json", contract, docx)
        self.assertFalse(result["passed"])
        self.assertTrue(any("ground pressure" in failure for failure in result["failures"]))


if __name__ == "__main__":
    unittest.main()
