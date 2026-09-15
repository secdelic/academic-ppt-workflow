from __future__ import annotations

import copy
import json
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest
import zipfile

from scripts.experiments.visual_fidelity_001 import (
    BASELINE, GENERATOR, ROOT, ZH_FIXTURE, arm_spec, canonical_json_hash, frozen_project_binding_issues, generate, normalized_tokens, package_snapshot, prepare,
)


class VisualFidelityExperimentTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory(prefix="vf001-tests-")
        cls.addClassCleanup(cls.temp.cleanup)
        cls.root = Path(cls.temp.name)
        cls.bundle = prepare(cls.root)
        cls.a_spec, cls.b_spec = arm_spec(cls.bundle, False), arm_spec(cls.bundle, True)
        cls.a = package_snapshot(generate(cls.a_spec, cls.root / "a"))
        cls.b = package_snapshot(generate(cls.b_spec, cls.root / "b"))
        cls.execution = json.loads((cls.root / "b/preview/experimental-execution.json").read_text())
        # Read-only stable renderer oracle; kept in OS temp, never production code.
        cls.old_script = cls.root / "stable_renderer.mjs"
        old = subprocess.run(["git", "show", f"{BASELINE}:scripts/generate_deck_pptxgen.mjs"], cwd=ROOT, capture_output=True, check=True).stdout
        cls.old_script.write_bytes(old)
        cls.old = package_snapshot(generate(cls.a_spec, cls.root / "old", cls.old_script))

    def compose(self, spec):
        module = (ROOT / "scripts/experiments/visual_fidelity_001.mjs").as_uri()
        code = f'import fs from "node:fs"; import {{prepareExperiment}} from {json.dumps(module)}; prepareExperiment(JSON.parse(fs.readFileSync(0,"utf8")));'
        return subprocess.run([shutil.which("node"), "--input-type=module", "-e", code], input=json.dumps(spec), capture_output=True, text=True, timeout=30)

    def assert_rejected(self, spec, text):
        result = self.compose(spec)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(text, result.stderr)

    def test_default_off_matches_stable_renderer(self):
        spec = copy.deepcopy(self.a_spec)
        spec.pop("experiment")
        actual = package_snapshot(generate(spec, self.root / "default"))
        self.assertEqual(actual, self.old)
        self.assertFalse((self.root / "default/preview/experimental-execution.json").exists())

    def test_disabled_matches_stable_renderer(self):
        self.assertEqual(self.a, self.old)

    def test_id_without_enabled_stays_off(self):
        spec = copy.deepcopy(self.b_spec)
        spec["experiment"].pop("enabled")
        actual = package_snapshot(generate(spec, self.root / "id-only"))
        self.assertEqual(actual, self.old)

    def test_enabled_with_wrong_id_fails_before_output(self):
        for i, value in enumerate([{"enabled": True}, {"enabled": True, "id": "other"}]):
            spec = copy.deepcopy(self.b_spec)
            spec["experiment"] = value
            with self.assertRaisesRegex(RuntimeError, "INVALID_VISUAL_FIDELITY_EXPERIMENT_ID"):
                generate(spec, self.root / f"wrong-{i}")
            self.assertFalse((self.root / f"wrong-{i}/deck.pptx").exists())

    def test_unknown_archetype_fails(self):
        spec = copy.deepcopy(self.b_spec)
        next(iter(spec["visual_execution_plan"].values()))["archetype"] = "INFERRED"
        self.assert_rejected(spec, "unknown archetype")

    def test_unknown_slide_fails(self):
        spec = copy.deepcopy(self.b_spec)
        spec["visual_execution_plan"]["UNKNOWN"] = {"archetype": "PROCESS_HORIZONTAL"}
        self.assert_rejected(spec, "Unknown experimental slide")

    def test_process_steps_and_only_frozen_edges(self):
        row = self.execution["slides"][0]
        objects = row["objects"]
        self.assertEqual([o["text"] for o in objects if o["kind"] == "text"], ["01 Intake", "02 Check", "03 Compose", "04 Archive"])
        self.assertEqual([o["edge"] for o in objects if "edge" in o], self.b_spec["slides"][0]["diagram_spec"]["edges"])
        self.assertGreater(len({o["bounds"]["h"] for o in objects if o["kind"] == "ellipse"}), 1)

    def test_process_rejects_unfrozen_order(self):
        spec = copy.deepcopy(self.b_spec)
        next(iter(spec["visual_execution_plan"].values()))["node_order"].reverse()
        self.assert_rejected(spec, "linear order")

    def test_hub_central_and_peripheral_nodes_undirected(self):
        row = self.execution["slides"][1]
        self.assertEqual(row["archetype_rendered"], "FRAMEWORK_HUB")
        self.assertEqual(sum(o["kind"] == "ellipse" for o in row["objects"]), 4)
        self.assertEqual(sum("edge" in o for o in row["objects"]), 3)
        self.assertTrue(all(not o["directed"] for o in row["objects"] if "edge" in o))
        self.assertEqual([o["id"] for o in row["objects"] if o.get("anchor")], ["node-core"])

    def test_system_has_only_explicit_edges_with_clipped_endpoints(self):
        row = self.execution["slides"][2]
        self.assertEqual([o["edge"] for o in row["objects"] if "edge" in o], self.b_spec["slides"][3]["diagram_spec"]["edges"])
        for edge in [o for o in row["objects"] if "edge" in o]:
            for key in ("source", "target"):
                node = next(o for o in row["objects"] if o["id"] == "node-" + edge["edge"][key])
                b, n = edge["bounds"], node["bounds"]
                # Neither end may terminate at the node center.
                center = (n["x"] + n["w"] / 2, n["y"] + n["h"] / 2)
                self.assertNotEqual(center, (b["x"], b["y"]))
                self.assertNotEqual(center, (b["x"] + b["w"], b["y"] + b["h"]))

    def test_system_rejects_unknown_endpoints(self):
        spec = copy.deepcopy(self.b_spec)
        spec["slides"][3]["diagram_spec"]["edges"][0]["target"] = "invented"
        self.assert_rejected(spec, "invalid/duplicate endpoints")

    def test_editorial_preserves_six_complete_statements(self):
        expected = [p for t in self.b_spec["slides"][4]["takeaways"] for p in t.splitlines()]
        row = self.execution["slides"][3]
        self.assertEqual([o["text"] for o in row["objects"] if o["kind"] == "text"], expected)
        self.assertEqual(len(expected), 6)
        for arm in (self.a, self.b):
            actual = [p for t in arm["slides"][4]["texts"] for p in t.splitlines()]
            self.assertTrue(all(actual.count(p) == 1 for p in expected))

    def test_scientific_text_mutation_rejected(self):
        spec = copy.deepcopy(self.b_spec)
        spec["slides"][0]["diagram_spec"]["nodes"][0]["label"] = "NEW SCIENTIFIC CLAIM"
        self.assert_rejected(spec, "node text must exactly preserve")

    def test_canonical_ir_and_hashes_unchanged(self):
        original = (self.bundle["staging"] / "deck_ir.json").read_bytes()
        for spec in (self.a_spec, self.b_spec):
            self.assertEqual(spec["deck_ir"]["canonical_hash"], self.bundle["ir"]["canonical_hash"])
            self.assertEqual(spec["slides"], self.bundle["canonical"]["slides"])
        self.compose(self.b_spec)
        self.assertEqual(original, (self.bundle["staging"] / "deck_ir.json").read_bytes())
        self.assertNotIn(b"visual_execution_plan", original)

    def test_source_citation_notes_unchanged(self):
        self.assertEqual(self.a["notes"], self.b["notes"])
        for a, b in zip(self.a_spec["slides"], self.b_spec["slides"], strict=True):
            for key in ("source_ids", "claim_ids", "short_source_label", "prohibited_overstatement", "speaker_notes"):
                self.assertEqual(a[key], b[key])

    def test_internal_geometry_inside_existing_frame(self):
        for row in self.execution["slides"]:
            frame = row["frame"]
            for o in row["objects"]:
                b = o["bounds"]
                self.assertGreaterEqual(b["x"], frame["x"])
                self.assertGreaterEqual(b["y"], frame["y"])
                self.assertLessEqual(b["x"] + b["w"], frame["x"] + frame["w"] + 1e-6)
                self.assertLessEqual(b["y"] + b["h"], frame["y"] + frame["h"] + 1e-6)
                if o["kind"] == "text":
                    self.assertGreaterEqual(o["font_size_pt"], 18)

    def test_no_footer_intrusion(self):
        footer = self.b_spec["layout_contract"]["zones"]["footer_exclusion"]["y"]
        for row in self.execution["slides"]:
            self.assertTrue(all(o["bounds"]["y"] + o["bounds"]["h"] < footer for o in row["objects"]))

    def test_no_off_slide(self):
        for slide in self.b["slides"]:
            for o in slide["objects"]:
                b = o["bounds"]
                self.assertGreaterEqual(b["x"], 0)
                self.assertGreaterEqual(b["y"], 0)
                self.assertLessEqual(b["x"] + b["w"], 13.334)
                self.assertLessEqual(b["y"] + b["h"], 7.501)

    def test_incompatible_frame_fails_closed(self):
        spec = copy.deepcopy(self.b_spec)
        next(o for o in spec["slides"][0]["planned_geometry"] if o["object_id"].endswith(":structure"))["bounds"]["h"] = 1
        self.assert_rejected(spec, "EXPERIMENTAL_ARCHETYPE_FRAME_INCOMPATIBLE")

    def test_no_rasterized_semantic_diagrams(self):
        for i in (0, 1, 3, 4):
            objects = self.b["slides"][i]["objects"]
            self.assertTrue(any(o["kind"] == "text" for o in objects))
            self.assertFalse(any(o["kind"] == "image" for o in objects))

    def test_non_target_renderers_unchanged(self):
        for part, digest in self.a["parts"].items():
            if part in ("ppt/slides/slide3.xml", "ppt/slides/slide6.xml") or part.startswith(("ppt/charts/", "ppt/embeddings/", "ppt/media/")):
                self.assertEqual(digest, self.b["parts"][part])

    def test_visible_text_invariant(self):
        self.assertEqual([s["tokens"] for s in self.a["slides"]], [s["tokens"] for s in self.b["slides"]])

    def test_text_comparison_detects_deletion_addition_number_change(self):
        original = normalized_tokens(["Keep 6 statements."])
        for changed in ("Keep statements.", "Keep 7 statements.", "Keep 6 statements. Proven."):
            self.assertNotEqual(original, normalized_tokens([changed]))
        self.assertEqual(original, normalized_tokens(["Keep", "6\nstatements."]))

    def test_targeted_count_and_zero_fallback(self):
        self.assertEqual(self.execution["targeted_slide_count"], 4)
        self.assertEqual(self.execution["rendered_slide_count"], 4)
        self.assertEqual(self.execution["generic_fallback_count"], 0)
        self.assertTrue(all(s["archetype_requested"] == s["archetype_rendered"] for s in self.execution["slides"]))

    def test_chart_cannot_be_targeted(self):
        spec = copy.deepcopy(self.b_spec)
        spec["visual_execution_plan"][spec["slides"][2]["slide_id"]] = {"archetype": "SYSTEM_MAP"}
        self.assert_rejected(spec, "scientific/template renderers are protected")


class CjkCalibrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.temp = tempfile.TemporaryDirectory(prefix="vf001-cjk-")
        cls.addClassCleanup(cls.temp.cleanup)
        cls.root = Path(cls.temp.name)
        cls.bundle = prepare(cls.root, ZH_FIXTURE)
        cls.spec = arm_spec(cls.bundle, True)
        cls.snapshot = package_snapshot(generate(cls.spec, cls.root / "cjk"))
        cls.execution = json.loads((cls.root / "cjk/preview/experimental-execution.json").read_text(encoding="utf-8"))

    def reject(self, spec, phrase):
        module = (ROOT / "scripts/experiments/visual_fidelity_001.mjs").as_uri()
        code = f'import fs from "node:fs"; import {{prepareExperiment}} from {json.dumps(module)}; prepareExperiment(JSON.parse(fs.readFileSync(0,"utf8")));'
        result = subprocess.run([shutil.which("node"), "--input-type=module", "-e", code], input=json.dumps(spec), capture_output=True, text=True, timeout=30)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn(phrase, result.stderr)

    def test_six_chinese_stages_equal_sizes_and_exact_order(self):
        row = self.execution["slides"][0]
        nodes = [o for o in row["objects"] if o["id"].startswith("node-")]
        self.assertEqual(len(nodes), 6)
        self.assertEqual(len({(o["bounds"]["w"], o["bounds"]["h"]) for o in nodes}), 1)
        self.assertEqual([o["edge"] for o in row["objects"] if "edge" in o], self.spec["slides"][0]["diagram_spec"]["edges"])

    def test_hub_three_equal_peripheral_nodes_no_causal_arrows(self):
        row = self.execution["slides"][1]
        peripheral = [o for o in row["objects"] if o.get("semantic_role") == "equal_member"]
        self.assertEqual(len(peripheral), 3)
        self.assertEqual(len({(o["bounds"]["w"],o["bounds"]["h"],o["tone"],o["line_width"]) for o in peripheral}), 1)
        self.assertTrue(all(not o["directed"] for o in row["objects"] if "edge" in o))

    def test_five_node_system_has_no_terminal_highlight(self):
        row = self.execution["slides"][2]
        nodes = [o for o in row["objects"] if o["id"].startswith("node-")]
        self.assertEqual(len(nodes), 5)
        self.assertTrue(all(not o["anchor"] and o["semantic_role"] == "equal_member" for o in nodes))
        self.assertEqual(len({(o["tone"],o["line_width"]) for o in nodes}),1)
        self.assertEqual([o["edge"] for o in row["objects"] if "edge" in o], self.spec["slides"][3]["diagram_spec"]["edges"])

    def test_six_full_chinese_statements_equal_typography(self):
        row = self.execution["slides"][3]
        texts = [o for o in row["objects"] if o["kind"] == "text"]
        expected = [t for group in self.spec["slides"][4]["takeaways"] for t in group.splitlines()]
        self.assertEqual([o["text"] for o in texts],expected)
        self.assertEqual(len(texts),6)
        self.assertTrue(all(o["font_size_pt"]==20 and not o["hero"] and not o["bold"] for o in texts))

    def test_oversized_chinese_case_is_expected_rejection(self):
        spec = copy.deepcopy(self.spec)
        negative = self.bundle["fixture"]["capacity_negative"]
        sentence = negative["paragraph"] * negative["repeat_per_statement"]
        spec["slides"][4]["takeaways"] = [sentence+"\n"+sentence]*3
        self.reject(spec, "EXPERIMENTAL_ARCHETYPE_FRAME_INCOMPATIBLE")

    def test_no_font_reduction_below_approved_minimum(self):
        spec = copy.deepcopy(self.spec)
        next(iter(spec["visual_execution_plan"].values()))["font_size_pt"] = 17
        self.reject(spec, "font below approved minimum")

    def test_plan_labels_and_directions_cannot_drift(self):
        for key in ("nodes", "edges"):
            spec = copy.deepcopy(self.spec)
            plan = next(iter(spec["visual_execution_plan"].values()))
            if key == "nodes": plan[key][0]["label"] = "未经批准的标签"
            else: plan[key][0]["directed"] = False
            self.reject(spec, "must match the frozen hashed plan")

    def test_missing_semantic_source_rejected(self):
        spec=copy.deepcopy(self.spec)
        next(iter(spec["visual_execution_plan"].values())).pop("source_locator")
        self.reject(spec,"explicit source")

    def test_crossing_an_unrelated_node_is_rejected(self):
        spec=copy.deepcopy(self.spec)
        plan=spec["visual_execution_plan"][spec["slides"][3]["slide_id"]]
        plan["positions"]={"a":[0.16,0.19],"b":[0.50,0.19],"c":[0.84,0.19],"d":[0.32,0.80],"e":[0.68,0.80]}
        self.reject(spec,"edge crosses node")

    def test_plan_hash_covers_roles_emphasis_and_direction(self):
        plan=self.spec["visual_execution_plan"]
        original=canonical_json_hash(plan)
        for field in ("emphasis","relationship_type","source_locator","central_id"):
            changed=copy.deepcopy(plan)
            next(iter(changed.values()))[field]="changed"
            self.assertNotEqual(original,canonical_json_hash(changed))

    def test_chinese_subscripts_slashes_and_abbreviations_preserved(self):
        text="\n".join(self.snapshot["slides"][1]["texts"])
        for value in ("rSO₂/PbtO₂","TCD/TCCD","cEEG/qEEG","共同主题"):
            self.assertIn(value,text)

    def test_preserved_full_excerpt_binding_does_not_require_claim_to_equal_title(self):
        slide={"slide_id":"SLD-SYNTHETIC","single_key_message":"原有主信息","source_ids":["SRC-SYNTHETIC"],"claim_ids":["CLM-SYNTHETIC"]}
        claim={"claim_id":"CLM-SYNTHETIC","source_id":"SRC-SYNTHETIC","claim_text":"这是登记的完整模拟摘录。","source_sha256":"hash"}
        manifest=[{"source_id":"SRC-SYNTHETIC","sha256":"hash"}]
        notes={"SLD-SYNTHETIC":"这是登记的完整模拟摘录。\n原有主信息"}
        self.assertEqual(frozen_project_binding_issues([slide],[copy.deepcopy(slide)],[claim],manifest,notes),[])
        changed=copy.deepcopy(slide);changed["single_key_message"]="新的科学主张"
        self.assertTrue(frozen_project_binding_issues([changed],[slide],[claim],manifest,notes))

    def test_excerpt_or_source_hash_mutation_is_rejected(self):
        slide={"slide_id":"SLD-SYNTHETIC","source_ids":["SRC-SYNTHETIC"],"claim_ids":["CLM-SYNTHETIC"]}
        claim={"claim_id":"CLM-SYNTHETIC","source_id":"SRC-SYNTHETIC","claim_text":"冻结摘录","source_sha256":"hash"}
        manifest=[{"source_id":"SRC-SYNTHETIC","sha256":"hash"}]
        self.assertTrue(frozen_project_binding_issues([slide],[slide],[claim],manifest,{"SLD-SYNTHETIC":"不匹配"}))
        claim["source_sha256"]="changed"
        self.assertTrue(frozen_project_binding_issues([slide],[slide],[claim],manifest,{"SLD-SYNTHETIC":"冻结摘录"}))


if __name__ == "__main__":
    unittest.main()
