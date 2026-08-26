from __future__ import annotations

import json
import shutil
import subprocess
import tempfile
import unittest
import zipfile
from pathlib import Path

from lxml import etree

from scripts.academic_ppt.fast_production import (
    build_keep_baseline_fingerprints,
    validate_keep_baseline_fingerprints,
)
from scripts.academic_ppt.keep_preservation import (
    NS,
    build_slide_semantic_manifest,
    load_keep_normalization_allowlist,
)


REPO_ROOT = Path(__file__).resolve().parents[2]
GENERATOR = REPO_ROOT / "tests" / "fixtures" / "synthetic" / "keep_semantics" / "generate_fixture.cjs"


def _rewrite(source: Path, output: Path, mutations: dict[str, callable]) -> Path:
    with zipfile.ZipFile(source, "r") as incoming, zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as outgoing:
        for info in incoming.infolist():
            payload = incoming.read(info.filename)
            if info.filename in mutations:
                payload = mutations[info.filename](payload)
            outgoing.writestr(info, payload)
    return output


def _xml_mutation(callback):
    def mutate(payload: bytes) -> bytes:
        root = etree.fromstring(payload)
        callback(root)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    return mutate


class KeepSemanticPreservationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._tmp = tempfile.TemporaryDirectory()
        cls.root = Path(cls._tmp.name)
        cls.baseline = cls.root / "synthetic_keep.pptx"
        completed = subprocess.run(
            ["node", str(GENERATOR), str(cls.baseline)],
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        if completed.returncode != 0:
            raise AssertionError(completed.stderr or completed.stdout)

    @classmethod
    def tearDownClass(cls) -> None:
        cls._tmp.cleanup()

    def _manifest(self, path: Path, slide: int) -> dict:
        return build_slide_semantic_manifest(path, slide)

    def _assert_semantic_changed(self, path: Path, slide: int) -> None:
        before = self._manifest(self.baseline, slide)
        after = self._manifest(path, slide)
        self.assertNotEqual(
            before["semantic_fingerprint_sha256"],
            after["semantic_fingerprint_sha256"],
        )

    def test_allowlist_is_explicit_and_fail_closed(self) -> None:
        allowlist = load_keep_normalization_allowlist()
        self.assertEqual(len(allowlist["rules"]), 11)
        self.assertTrue(all(row["semantic_impact"] == "NONE" for row in allowlist["rules"]))
        self.assertFalse(any("*" in row["field"] for row in allowlist["rules"]))

    def test_com_regression_runner_uses_real_powerpoint_saveas(self) -> None:
        runner = (REPO_ROOT / "scripts" / "ci" / "run_keep_semantic_com_regression.py").read_text(encoding="utf-8")
        saveas = (REPO_ROOT / "scripts" / "powerpoint_saveas_copy.ps1").read_text(encoding="utf-8")
        self.assertIn("SECOND_DEVICE_COM_KEEP_TABLE_NORMALIZATION", runner)
        self.assertIn("Presentations.Open", saveas)
        self.assertIn("SaveAs", saveas)
        self.assertNotIn("mock", saveas.casefold())

    def test_powerpoint_table_serialization_normalization_matches(self) -> None:
        def normalize_slide(root: etree._Element) -> None:
            table_shape = root.find(".//p:graphicFrame", namespaces=NS)
            identity = table_shape.find(".//p:cNvPr", namespaces=NS)
            identity.set("id", str(int(identity.get("id")) + 1))
            for cell in root.xpath(".//a:tcPr", namespaces=NS):
                cell.attrib.pop("marT", None)
                cell.attrib.pop("marB", None)
            for col_index, column in enumerate(root.xpath(".//a:tblGrid/a:gridCol", namespaces=NS), 1):
                ext_list = etree.SubElement(column, f"{{{NS['a']}}}extLst")
                ext = etree.SubElement(ext_list, f"{{{NS['a']}}}ext")
                ext.set("uri", "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}")
                col_id = etree.SubElement(ext, "{http://schemas.microsoft.com/office/drawing/2014/main}colId")
                col_id.set("val", str(1000 + col_index))
            for row_index, row in enumerate(root.xpath(".//a:tr", namespaces=NS), 1):
                ext_list = etree.SubElement(row, f"{{{NS['a']}}}extLst")
                ext = etree.SubElement(ext_list, f"{{{NS['a']}}}ext")
                ext.set("uri", "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}")
                row_id = etree.SubElement(ext, "{http://schemas.microsoft.com/office/drawing/2014/main}rowId")
                row_id.set("val", str(2000 + row_index))

        normalized = _rewrite(
            self.baseline,
            self.root / "normalized.pptx",
            {"ppt/slides/slide2.xml": _xml_mutation(normalize_slide)},
        )
        before, after = self._manifest(self.baseline, 2), self._manifest(normalized, 2)
        self.assertEqual(before["semantic_fingerprint_sha256"], after["semantic_fingerprint_sha256"])

    def test_relationship_id_renumbering_with_same_target_matches(self) -> None:
        with zipfile.ZipFile(self.baseline) as package:
            slide = etree.fromstring(package.read("ppt/slides/slide3.xml"))
            old = slide.find(".//a:blip", namespaces=NS).get(f"{{{NS['r']}}}embed")

        def change_slide(root: etree._Element) -> None:
            root.find(".//a:blip", namespaces=NS).set(f"{{{NS['r']}}}embed", "rId99")

        def change_rels(root: etree._Element) -> None:
            for row in root:
                if row.get("Id") == old:
                    row.set("Id", "rId99")

        updated = _rewrite(
            self.baseline,
            self.root / "relationship_renumbered.pptx",
            {
                "ppt/slides/slide3.xml": _xml_mutation(change_slide),
                "ppt/slides/_rels/slide3.xml.rels": _xml_mutation(change_rels),
            },
        )
        self.assertEqual(
            self._manifest(self.baseline, 3)["semantic_fingerprint_sha256"],
            self._manifest(updated, 3)["semantic_fingerprint_sha256"],
        )

    def test_registered_slide_number_value_is_dynamic_but_format_is_not(self) -> None:
        def make_slide_number(value: str, *, font_size: str | None = None):
            def mutate(root: etree._Element) -> None:
                shape = root.find(".//p:sp", namespaces=NS)
                non_visual = shape.find("./p:nvSpPr/p:nvPr", namespaces=NS)
                placeholder = etree.SubElement(non_visual, f"{{{NS['p']}}}ph")
                placeholder.set("type", "sldNum")
                shape.find(".//a:t", namespaces=NS).text = value
                if font_size is not None:
                    shape.find(".//a:rPr", namespaces=NS).set("sz", font_size)

            return mutate

        before = _rewrite(
            self.baseline,
            self.root / "slide_number_before.pptx",
            {"ppt/slides/slide1.xml": _xml_mutation(make_slide_number("1"))},
        )
        after = _rewrite(
            self.baseline,
            self.root / "slide_number_after.pptx",
            {"ppt/slides/slide1.xml": _xml_mutation(make_slide_number("9"))},
        )
        changed_format = _rewrite(
            self.baseline,
            self.root / "slide_number_format_changed.pptx",
            {"ppt/slides/slide1.xml": _xml_mutation(make_slide_number("9", font_size="3600"))},
        )
        before_manifest = self._manifest(before, 1)
        self.assertEqual(
            before_manifest["semantic_fingerprint_sha256"],
            self._manifest(after, 1)["semantic_fingerprint_sha256"],
        )
        self.assertNotEqual(
            before_manifest["semantic_fingerprint_sha256"],
            self._manifest(changed_format, 1)["semantic_fingerprint_sha256"],
        )

    def test_negative_table_and_shape_mutations_fail(self) -> None:
        mutations = {
            "cell_text": lambda root: root.find(".//a:t", namespaces=NS).__setattr__("text", "Changed"),
            "font_size": lambda root: root.find(".//a:rPr", namespaces=NS).set("sz", "2800"),
            "cell_fill": lambda root: root.find(".//a:tcPr/a:solidFill/a:srgbClr", namespaces=NS).set("val", "FF0000"),
            "border": lambda root: root.find(".//a:tcPr/a:lnL", namespaces=NS).set("w", "25400"),
            "column_width": lambda root: root.find(".//a:tblGrid/a:gridCol", namespaces=NS).set("w", "3000000"),
            "row_height": lambda root: root.find(".//a:tr", namespaces=NS).set("h", "900000"),
            "move_table": lambda root: root.find(".//p:graphicFrame/p:xfrm/a:off", namespaces=NS).set("x", "2000000"),
        }
        for name, callback in mutations.items():
            with self.subTest(name=name):
                output = _rewrite(
                    self.baseline,
                    self.root / f"negative_{name}.pptx",
                    {"ppt/slides/slide2.xml": _xml_mutation(callback)},
                )
                self._assert_semantic_changed(output, 2)

    def test_negative_z_order_image_master_layout_and_chart_data_fail(self) -> None:
        def swap_shapes(root: etree._Element) -> None:
            tree = root.find(".//p:spTree", namespaces=NS)
            objects = [child for child in tree if etree.QName(child).localname in {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}]
            first_index, second_index = tree.index(objects[1]), tree.index(objects[2])
            tree.remove(objects[2]); tree.insert(first_index, objects[2])
            tree.remove(objects[1]); tree.insert(second_index, objects[1])

        z_order = _rewrite(
            self.baseline,
            self.root / "negative_z_order.pptx",
            {"ppt/slides/slide4.xml": _xml_mutation(swap_shapes)},
        )
        self._assert_semantic_changed(z_order, 4)

        with zipfile.ZipFile(self.baseline) as package:
            image_part = self._manifest(self.baseline, 3)["presentation_identity"]["media_part_names_diagnostic"][0]
            chart_part = next(name for name in package.namelist() if name.startswith("ppt/charts/chart") and name.endswith(".xml"))

        image = _rewrite(
            self.baseline,
            self.root / "negative_image.pptx",
            {image_part: lambda payload: payload + b"SEMANTIC_IMAGE_CHANGE"},
        )
        self._assert_semantic_changed(image, 3)

        def change_layout(root: etree._Element) -> None:
            for row in root:
                if str(row.get("Type", "")).endswith("/slideLayout"):
                    row.set("Target", "../slideLayouts/slideLayout999.xml")

        layout = _rewrite(
            self.baseline,
            self.root / "negative_layout.pptx",
            {"ppt/slides/_rels/slide1.xml.rels": _xml_mutation(change_layout)},
        )
        self.assertNotEqual(
            self._manifest(self.baseline, 1)["presentation_identity"],
            self._manifest(layout, 1)["presentation_identity"],
        )

        def change_chart(root: etree._Element) -> None:
            value = root.find(".//c:v", namespaces=NS)
            value.text = "999"

        chart = _rewrite(
            self.baseline,
            self.root / "negative_chart.pptx",
            {chart_part: _xml_mutation(change_chart)},
        )
        self._assert_semantic_changed(chart, 6)

    def test_three_tier_gate_requires_render_when_only_raw_differs(self) -> None:
        def omit_explicit_default_margins(root: etree._Element) -> None:
            for cell in root.xpath(".//a:tcPr", namespaces=NS):
                cell.attrib.pop("marT", None)
                cell.attrib.pop("marB", None)

        normalized = _rewrite(
            self.baseline,
            self.root / "raw_only_difference.pptx",
            {"ppt/slides/slide2.xml": _xml_mutation(omit_explicit_default_margins)},
        )
        ids = [f"SYN-{index}" for index in range(1, 7)]
        plan = {
            "source_slide_ids": ids,
            "expected_final_order": ids,
            "operations": [{"action": "KEEP", "slide_id": value} for value in ids],
        }
        baseline = build_keep_baseline_fingerprints(source_pptx=self.baseline, operation_plan=plan)
        first = validate_keep_baseline_fingerprints(
            baseline=baseline,
            updated_pptx=normalized,
            expected_final_order=ids,
            operation_plan=plan,
        )
        self.assertEqual(first["status"], "RENDER_IDENTITY_REQUIRED")
        self.assertEqual(first["render_required_slide_ids"], ["SYN-2"])
        final = validate_keep_baseline_fingerprints(
            baseline=baseline,
            updated_pptx=normalized,
            expected_final_order=ids,
            operation_plan=plan,
            render_identity={"SYN-2": {"status": "PASS"}},
        )
        self.assertEqual(final["status"], "PASS")


if __name__ == "__main__":
    unittest.main()
