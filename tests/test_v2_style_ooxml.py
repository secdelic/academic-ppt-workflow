from __future__ import annotations

import base64
import hashlib
import json
import tempfile
import unittest
import zipfile
from pathlib import Path

import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

from academic_ppt.ooxml_qa import (
    OOXMLQAError,
    inspect_package,
    inspect_pptx_ooxml,
    validate_ir_identity_map,
)
from academic_ppt.style_reference import (
    ReferenceStyleError,
    parse_reference_style,
)


PRIVATE_SENTINELS = (
    "SENTINEL_PRIVATE_PATIENT_X",
    "SENTINEL_PRIVATE_ALT",
    "SENTINEL_PRIVATE_NOTE",
    "SENTINEL_PRIVATE_CHART_VALUE",
    "SENTINEL_PRIVATE_EXTERNAL_TARGET",
)

REL_NS = (
    "http://schemas.openxmlformats.org/package/2006/relationships"
)
DOC_REL_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)


def _relationships(rows: list[tuple[str, str, str, str | None]]) -> str:
    elements = []
    for relationship_id, relationship_type, target, target_mode in rows:
        target_mode_attribute = (
            f' TargetMode="{target_mode}"' if target_mode else ""
        )
        elements.append(
            f'<Relationship Id="{relationship_id}" Type="{relationship_type}" '
            f'Target="{target}"{target_mode_attribute}/>'
        )
    return (
        f'<?xml version="1.0" encoding="UTF-8"?>'
        f'<Relationships xmlns="{REL_NS}">'
        + "".join(elements)
        + "</Relationships>"
    )


def _write_minimal_pptx(path: Path, *, broken_media: bool = False) -> None:
    presentation_type = (
        "application/vnd.openxmlformats-officedocument.presentationml."
        "presentation.main+xml"
    )
    slide_type = (
        "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    )
    layout_type = (
        "application/vnd.openxmlformats-officedocument.presentationml."
        "slideLayout+xml"
    )
    master_type = (
        "application/vnd.openxmlformats-officedocument.presentationml."
        "slideMaster+xml"
    )
    notes_type = (
        "application/vnd.openxmlformats-officedocument.presentationml."
        "notesSlide+xml"
    )
    theme_type = "application/vnd.openxmlformats-officedocument.theme+xml"
    chart_type = (
        "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    )
    chart_colors_type = (
        "application/vnd.ms-office.chartcolorstyle+xml"
    )
    content_types = f"""<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="{presentation_type}"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="{slide_type}"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="{layout_type}"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="{master_type}"/>
  <Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="{notes_type}"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="{theme_type}"/>
  <Override PartName="/ppt/charts/chart1.xml" ContentType="{chart_type}"/>
  <Override PartName="/ppt/charts/colors1.xml" ContentType="{chart_colors_type}"/>
</Types>"""

    presentation = f"""<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:r="{DOC_REL_NS}" showSpecialPlsOnTitleSld="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId2"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
</p:presentation>"""

    slide = f"""<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
 xmlns:r="{DOC_REL_NS}">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Title" descr="SENTINEL_PRIVATE_ALT"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="500000" y="300000"/><a:ext cx="11000000" cy="800000"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr lIns="91440" rIns="91440" tIns="45720" bIns="45720"/><a:lstStyle/><a:p><a:r><a:t>SENTINEL_PRIVATE_PATIENT_X</a:t></a:r></a:p></p:txBody>
    </p:sp>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="3" name="Private image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
      <p:spPr><a:xfrm><a:off x="500000" y="1500000"/><a:ext cx="2000000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
    </p:pic>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="4" name="Private chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="3000000" y="1500000"/><a:ext cx="4000000" cy="2500000"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId4"/></a:graphicData></a:graphic>
    </p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>"""

    layout = """<?xml version="1.0" encoding="UTF-8"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" type="obj" preserve="1">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name=""/><p:cNvSpPr/><p:nvPr><p:ph type="title" idx="0"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="500000" y="300000"/><a:ext cx="11000000" cy="800000"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr lIns="91440" rIns="91440" tIns="45720" bIns="45720"/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="3" name=""/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="500000" y="1400000"/><a:ext cx="11000000" cy="4700000"/></a:xfrm></p:spPr>
      <p:txBody><a:bodyPr lIns="91440" rIns="91440" tIns="45720" bIns="45720"/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:hf dt="0" ftr="1" sldNum="1"/>
</p:sldLayout>"""

    master = """<?xml version="1.0" encoding="UTF-8"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name=""/><p:cNvSpPr/><p:nvPr><p:ph type="ftr"/></p:nvPr></p:nvSpPr>
      <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
  <p:hf dt="0" ftr="1" sldNum="1"/>
</p:sldMaster>"""

    theme = """<?xml version="1.0" encoding="UTF-8"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <a:themeElements>
    <a:clrScheme name="Synthetic">
      <a:dk1><a:srgbClr val="1F2937"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:accent1><a:srgbClr val="0072B2"/></a:accent1>
      <a:accent2><a:srgbClr val="D55E00"/></a:accent2>
    </a:clrScheme>
    <a:fontScheme name="Synthetic">
      <a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface="Arial"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface="Arial"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Synthetic"/>
  </a:themeElements>
</a:theme>"""

    notes = """<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name=""/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>SENTINEL_PRIVATE_NOTE</a:t></a:r></a:p><a:p><a:r><a:t>[Sources]</a:t></a:r></a:p><a:p><a:r><a:t>SRC-SYNTH-001</a:t></a:r></a:p></p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:notes>"""

    chart = """<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:numCache>
    <c:pt idx="0"><c:v>SENTINEL_PRIVATE_CHART_VALUE</c:v></c:pt>
  </c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart>
</c:chartSpace>"""
    chart_colors = """<?xml version="1.0" encoding="UTF-8"?>
<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <cs:color><a:schemeClr val="accent1"/></cs:color>
  <cs:color><a:srgbClr val="CC6677"/></cs:color>
</cs:colorStyle>"""
    table_styles = """<?xml version="1.0" encoding="UTF-8"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 def="{5940675A-B579-460E-94D1-54222C63F5DA}">
  <a:tblStyle styleId="{5940675A-B579-460E-94D1-54222C63F5DA}"/>
</a:tblStyleLst>"""

    root_rels = _relationships(
        [
            (
                "rId1",
                DOC_REL_NS + "/officeDocument",
                "ppt/presentation.xml",
                None,
            )
        ]
    )
    presentation_rels = _relationships(
        [
            ("rId1", DOC_REL_NS + "/slide", "slides/slide1.xml", None),
            (
                "rId2",
                DOC_REL_NS + "/slideMaster",
                "slideMasters/slideMaster1.xml",
                None,
            ),
        ]
    )
    media_target = (
        "../media/missing.png" if broken_media else "../media/image1.png"
    )
    slide_rels = _relationships(
        [
            (
                "rId1",
                DOC_REL_NS + "/slideLayout",
                "../slideLayouts/slideLayout1.xml",
                None,
            ),
            (
                "rId2",
                DOC_REL_NS + "/notesSlide",
                "../notesSlides/notesSlide1.xml",
                None,
            ),
            ("rId3", DOC_REL_NS + "/image", media_target, None),
            ("rId4", DOC_REL_NS + "/chart", "../charts/chart1.xml", None),
            (
                "rId5",
                DOC_REL_NS + "/hyperlink",
                "https://example.invalid/SENTINEL_PRIVATE_EXTERNAL_TARGET",
                "External",
            ),
        ]
    )
    layout_rels = _relationships(
        [
            (
                "rId1",
                DOC_REL_NS + "/slideMaster",
                "../slideMasters/slideMaster1.xml",
                None,
            )
        ]
    )
    master_rels = _relationships(
        [
            (
                "rId1",
                DOC_REL_NS + "/slideLayout",
                "../slideLayouts/slideLayout1.xml",
                None,
            ),
            ("rId2", DOC_REL_NS + "/theme", "../theme/theme1.xml", None),
        ]
    )
    notes_rels = _relationships(
        [
            ("rId1", DOC_REL_NS + "/slide", "../slides/slide1.xml", None),
        ]
    )
    one_pixel_png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk"
        "+A8AAQUBAScY42YAAAAASUVORK5CYII="
    )

    members: dict[str, str | bytes] = {
        "[Content_Types].xml": content_types,
        "_rels/.rels": root_rels,
        "ppt/presentation.xml": presentation,
        "ppt/_rels/presentation.xml.rels": presentation_rels,
        "ppt/slides/slide1.xml": slide,
        "ppt/slides/_rels/slide1.xml.rels": slide_rels,
        "ppt/slideLayouts/slideLayout1.xml": layout,
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": layout_rels,
        "ppt/slideMasters/slideMaster1.xml": master,
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": master_rels,
        "ppt/theme/theme1.xml": theme,
        "ppt/notesSlides/notesSlide1.xml": notes,
        "ppt/notesSlides/_rels/notesSlide1.xml.rels": notes_rels,
        "ppt/charts/chart1.xml": chart,
        "ppt/charts/colors1.xml": chart_colors,
        "ppt/tableStyles.xml": table_styles,
        "ppt/media/image1.png": one_pixel_png,
    }
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as package:
        for member, content in members.items():
            package.writestr(
                member,
                content.encode("utf-8") if isinstance(content, str) else content,
            )


class ReferenceStyleTests(unittest.TestCase):
    def test_style_only_is_read_only_and_does_not_leak_content(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "private_reference.pptx"
            _write_minimal_pptx(path)
            before = hashlib.sha256(path.read_bytes()).hexdigest()
            profile = parse_reference_style(path)
            after = hashlib.sha256(path.read_bytes()).hexdigest()
            serialized = json.dumps(profile, ensure_ascii=False)

            self.assertEqual(before, after)
            self.assertEqual(profile["source"]["sha256"], before)
            self.assertTrue(profile["source"]["read_only_verified"])
            self.assertEqual(
                profile["slide_dimensions"]["aspect_ratio"], 1.7778
            )
            self.assertEqual(
                profile["theme_fonts"]["major"]["latin"], "Aptos Display"
            )
            self.assertTrue(
                any(
                    item["role"] == "title_and_content"
                    for item in profile["common_page_roles"]
                )
            )
            self.assertTrue(profile["image_aspect_ratio_patterns"])
            self.assertTrue(
                any(
                    token["value"] == "CC6677"
                    for token in profile["chart_color_tokens"]["tokens"]
                )
            )
            self.assertIn(
                "{5940675A-B579-460E-94D1-54222C63F5DA}",
                profile["table_styles"]["style_ids"],
            )
            self.assertEqual(
                profile["external_relationship_warnings"][0]["action"],
                "NOT_FOLLOWED",
            )
            for sentinel in PRIVATE_SENTINELS:
                self.assertNotIn(sentinel, serialized)

    def test_all_modes_are_explicit_and_content_reference_is_hash_only(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "reference.potx"
            _write_minimal_pptx(path)
            for mode in (
                "style-only",
                "template-fill",
                "content-reference",
                "protected-reference",
            ):
                profile = parse_reference_style(path, mode=mode)
                self.assertEqual(profile["reference_mode"], mode)
                serialized = json.dumps(profile, ensure_ascii=False)
                for sentinel in PRIVATE_SENTINELS:
                    self.assertNotIn(sentinel, serialized)
            content_profile = parse_reference_style(
                path, mode="content-reference"
            )
            self.assertFalse(
                content_profile["content_reference_manifest"]["content_returned"]
            )

    def test_corrupt_reference_fails_closed(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "corrupt.pptx"
            path.write_bytes(b"not a zip")
            with self.assertRaises(ReferenceStyleError):
                parse_reference_style(path)


class OOXMLQATests(unittest.TestCase):
    def test_structural_qa_chains_sources_and_privacy(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "valid.pptx"
            _write_minimal_pptx(path)
            before = hashlib.sha256(path.read_bytes()).hexdigest()
            report = inspect_package(
                path, known_source_ids={"SRC-SYNTH-001"}
            )
            after = hashlib.sha256(path.read_bytes()).hexdigest()
            serialized = json.dumps(report, ensure_ascii=False)

            self.assertEqual(before, after)
            self.assertTrue(report["valid"])
            self.assertEqual(report["status"], "WARN")
            self.assertTrue(
                report["slide_relationship_chains"][0]["chain_valid"]
            )
            self.assertTrue(
                report["notes_source_checks"][0]["has_sources_marker"]
            )
            self.assertEqual(
                report["notes_source_checks"][0]["source_id_token_count"], 1
            )
            self.assertEqual(
                report["notes_source_checks"][0]["unknown_source_id_count"], 0
            )
            self.assertEqual(len(report["canonical_slide_hashes"]), 1)
            self.assertEqual(
                report["security_flags"]["external_relationship_count"], 1
            )
            self.assertFalse(
                report["security_flags"]["followed_external_relationships"]
            )
            self.assertTrue(
                all(
                    issue["severity"] in {"ERROR", "WARNING"}
                    for issue in report["issues"]
                )
            )
            for sentinel in PRIVATE_SENTINELS:
                self.assertNotIn(sentinel, serialized)

    def test_broken_relationship_is_detected(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "broken.pptx"
            _write_minimal_pptx(path, broken_media=True)
            report = inspect_pptx_ooxml(path)
            codes = {item["code"] for item in report["errors"]}
            self.assertFalse(report["valid"])
            self.assertEqual(report["status"], "FAIL")
            self.assertIn("DANGLING_RELATIONSHIP", codes)
            self.assertEqual(report["relationship_summary"]["dangling_count"], 1)

    def test_corrupt_package_fails_closed(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "corrupt.pptx"
            path.write_bytes(b"not a zip")
            with self.assertRaises(OOXMLQAError):
                inspect_pptx_ooxml(path)

    def test_ir_identity_map_uses_stable_id_not_page_number(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "valid.pptx"
            _write_minimal_pptx(path)
            report = inspect_pptx_ooxml(path)
            result = validate_ir_identity_map(
                report,
                {
                    "slides": [
                        {
                            "slide_id": "results-primary-outcome-v1",
                            "slide_revision": 2,
                            "pptx_slide_part": "ppt/slides/slide1.xml",
                            "content_hash": "synthetic-ir-hash",
                        }
                    ]
                },
            )
            self.assertTrue(result["consistent"])
            self.assertEqual(
                result["identity_map"][0]["slide_id"],
                "results-primary-outcome-v1",
            )
            self.assertEqual(
                result["identity_map"][0]["pptx_slide_part"],
                "ppt/slides/slide1.xml",
            )


if __name__ == "__main__":
    unittest.main()
