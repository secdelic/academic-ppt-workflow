"""Hash-bound visual process checks. Disabled calls perform no I/O.

Enforce mode exposes blocking findings to the production delivery gate.
Human aesthetic judgments are never inferred from renderer metadata.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
import csv
import hashlib
import json
from pathlib import Path
import zipfile
from xml.etree import ElementTree as ET

from .ooxml_qa import NS, inspect_pptx_ooxml

ROOT = Path(__file__).resolve().parents[2]
CONTRACTS = ROOT / "config/visual_quality"
VERSION = "visual-quality-floor/1"
DECISIONS = {
    "NO_ILLUSTRATION_NEEDED", "ILLUSTRATION_RECOMMENDED", "SOURCE_FIGURE_REQUIRED",
    "APPROVED_ASSET_SELECTED", "AI_CONCEPTUAL_CANDIDATE_POSSIBLE", "HUMAN_DECISION_REQUIRED",
}
ART_FIELDS = (
    "palette", "background_strategy", "typography_personality", "hero_style",
    "card_style", "annotation_style", "divider_style", "density_pattern",
    "signature_components", "approved_assets",
)
OUTPUTS = (
    "visual_quality_report.json", "visual_quality_report.md", "slide_quality_records.json",
    "deck_quality_summary.json", "art_direction_execution_report.json",
    "illustration_decision_report.json", "human_review_queue.csv", "quality_floor_execution_log.md",
)


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def digest(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False).encode()).hexdigest()


def reference(path, pointer=None):
    result = {"path": str(Path(path).resolve()), "sha256": sha256(path)}
    if pointer is not None:
        result["pointer"] = pointer
    return result


def pointer_value(value, pointer):
    for key in pointer.strip("/").split("/") if pointer else []:
        key = key.replace("~1", "/").replace("~0", "~")
        value = value[int(key)] if isinstance(value, list) else value[key]
    return value


def load_authority():
    lock = read_json(CONTRACTS / "authority.json")
    for name, expected in lock["contract_sha256"].items():
        if sha256(CONTRACTS / name) != expected:
            raise ValueError("QUALITY_FLOOR_DESIGN_CONTRACT_CHANGE_REQUIRED: " + name)
    return {"lock": lock, **{name: read_json(CONTRACTS / name)
                            for name in lock["contract_sha256"]}}


def normalized_family(family):
    if not isinstance(family, str) or not family.strip():
        return "UNRESOLVED"
    if family.upper().startswith(("GENERIC_CARD", "GENERIC_PANEL")):
        return "GENERIC_CARD_GRID"
    return family


def longest_run(values, target=None):
    last, length, maximum = None, 0, 0
    for value in values:
        length = length + 1 if value == last else 1
        if target is None or value == target:
            maximum = max(maximum, length)
        last = value
    return maximum


def structure_groups(manifests, width, height):
    """Macro geometry, ignoring labels/colors/IDs; 2.5% canvas tolerance.

    This describes a repetition candidate, never a semantic or aesthetic verdict.
    Compare each page to the cluster representative, avoiding transitive drift.
    """
    representatives, groups = [], []
    for slide in manifests:
        objects = []
        for obj in slide["objects"]:
            g = obj["geometry"]
            if any(g.get(k) is None for k in ("x_emu", "y_emu", "width_emu", "height_emu")):
                continue
            objects.append((obj["kind"], g["x_emu"] / width, g["y_emu"] / height,
                            g["width_emu"] / width, g["height_emu"] / height))
        objects.sort()
        match = None
        if objects:
            for i, rep in enumerate(representatives):
                if len(rep) == len(objects) and all(a[0] == b[0] and
                        max(abs(x-y) for x, y in zip(a[1:], b[1:])) <= .025
                        for a, b in zip(rep, objects)):
                    match = i
                    break
        if match is None:
            match = len(representatives)
            representatives.append(objects)
        groups.append(f"STRUCTURE_{match + 1:03d}")
    return groups


def evaluate(manifest, *, mode="off"):
    """Return a report in explicit shadow mode; default is a side-effect-free no-op."""
    if mode == "off":
        return None
    if mode not in ("shadow", "enforce"):
        raise ValueError("Expected off, shadow or enforce mode")
    authority = load_authority()
    matrix = {c["check_id"]: c for c in authority["quality_floor_QA_matrix.json"]["checks"]}
    taxonomy = authority["page_role_and_composition_taxonomy.json"]
    roles = {r["page_role"] for r in taxonomy["roles"]}
    safety = authority["ai_illustration_safety_contract.json"]
    findings, checked = [], {}

    def issue(check, code, basis, sid=None, result="FAIL"):
        findings.append({"check_id": check, "severity": matrix[check]["severity"],
                         "code": code, "slide_id": sid, "result": result, "basis": basis})

    def verify(ref):
        if not isinstance(ref, dict) or not ref.get("path") or not ref.get("sha256"):
            return False
        p = Path(ref["path"])
        if not p.is_file() or sha256(p) != ref["sha256"]:
            return False
        checked[str(p.resolve())] = ref["sha256"]
        return True

    def fact(ref):
        if not verify(ref):
            return None
        try:
            return pointer_value(read_json(ref["path"]), ref.get("pointer", ""))
        except (KeyError, IndexError, ValueError, TypeError):
            return None

    def qa(name, check, code):
        receipt = manifest.get("qa", {}).get(name, {})
        status = fact(receipt.get("status"))
        bound = fact(receipt.get("deck_binding")) == deck_sha
        if status not in ("PASS", "PASS_WITH_INHERITED_SYMBOL_FALLBACK") or not bound:
            issue(check, code, f"{name}: {status if bound else 'NOT_ASSESSED (missing/stale deck binding)'}")
            return "FAIL" if status == "FAIL" and bound else "NOT_ASSESSED"
        return "PASS"

    deck = manifest.get("deck", {})
    if not verify(deck):
        raise ValueError("INFORMATION_REQUIRED: missing or changed PPTX; cannot establish page denominator")
    deck_sha = deck["sha256"]
    package = inspect_pptx_ooxml(deck["path"])
    if not package["valid"]:
        issue("Q08", "OOXML_QA_FAILURE", json.dumps(package["errors"]))
    count = len(package["slide_order"])
    if not count:
        issue("Q01", "EMPTY_DECK", "Actual presentation has no pages")
    sources = manifest.get("scientific_inputs", [])
    if not sources or not all([verify(ref) for ref in sources]):
        issue("Q02", "SCIENTIFIC_INPUT_EVIDENCE_MISSING", "Frozen scientific input hashes are required")
    scientific = qa("scientific", "Q02", "SCIENTIFIC_QA_FAILURE")
    geometry = qa("geometry", "Q07", "GEOMETRY_QA_FAILURE")
    geometry_pages = fact(manifest.get("qa", {}).get("geometry", {}).get("pages"))
    if (not isinstance(geometry_pages, list) or
            [p.get("slide_index") for p in geometry_pages] != list(range(1, count+1))):
        issue("Q07", "GEOMETRY_PAGE_COVERAGE_INCOMPLETE", "Geometry evidence must cover every actual page in order")
        geometry = "NOT_ASSESSED"
    typography = qa("typography", "Q07", "TYPOGRAPHY_MINIMUM_NOT_VERIFIED")
    qa("semantic", "Q03", "SEMANTIC_QA_FAILURE")

    render = manifest.get("render", {})
    render_bound = fact(render.get("deck_binding")) == deck_sha
    previews = render.get("previews", [])
    previews_by_index = {p.get("slide_index"): p for p in previews}
    from PIL import Image, UnidentifiedImageError
    from pypdf import PdfReader

    def readable_image(ref):
        if not verify(ref):
            return False
        try:
            with Image.open(ref["path"]) as img:
                img.verify()
        except (OSError, ValueError, UnidentifiedImageError):
            return False
        return True

    if (not render_bound or len(previews) != count or
            set(previews_by_index) != set(range(1, count+1)) or
            not all([readable_image(p) for p in previews])):
        issue("Q08", "SLIDE_RENDER_MISSING_OR_STALE", "Every actual page needs an indexed, hash-bound readable preview")
    sheet = render.get("contact_sheet", {})
    if (not readable_image(sheet) or not render_bound or
            sheet.get("slide_indexes") != list(range(1, count+1))):
        issue("Q08", "CONTACT_SHEET_MISSING_OR_STALE", "An exact-deck full contact sheet is required")
    pdf = render.get("pdf", {})
    if not verify(pdf) or not render_bound:
        issue("Q08", "PDF_RENDER_MISSING_OR_STALE", "Existing rendered PDF and binding required")
    else:
        if len(PdfReader(pdf["path"]).pages) != count:
            issue("Q08", "PDF_PAGE_COUNT_MISMATCH", "PDF page count differs from actual deck")

    raw_records = manifest.get("records", [])
    ids = [r.get("slide_id") for r in raw_records]
    indexes = [r.get("slide_index") for r in raw_records]
    if (len(raw_records) != count or len(set(ids)) != len(ids) or
            len(set(indexes)) != len(indexes) or set(indexes) != set(range(1, count+1))):
        issue("Q01", "SLIDE_RECORD_COVERAGE_INVALID", "Records must match actual pages exactly, with unique IDs and indexes")
    by_index = {r.get("slide_index"): r for r in raw_records}
    observations = manifest.get("observations", [])
    if len({o.get("slide_index") for o in observations}) != len(observations):
        issue("Q01", "DUPLICATE_OBSERVATION", "Ambiguous actual composition observations")
    obs_by_index = {o.get("slide_index"): o for o in observations}
    records = []
    with zipfile.ZipFile(deck["path"]) as archive:
        size = ET.fromstring(archive.read("ppt/presentation.xml")).find("p:sldSz", NS)
        structure = structure_groups(package["object_manifest"], int(size.get("cx")), int(size.get("cy")))
        for page in package["slide_order"]:
            i = page["slide_index"]
            r = by_index.get(i, {})
            sid = r.get("slide_id") or f"UNRESOLVED_PAGE_{i}"
            if not r.get("slide_id"):
                issue("Q01", "MISSING_SLIDE_ID", "No valid identity record for actual page", sid)
            role = r.get("page_role")
            if not isinstance(role, str) or role not in roles:
                issue("Q01", "PAGE_ROLE_UNRESOLVED", "Exactly one source-bound primary page role required", sid)
            for field in authority["visual_quality_contract.json"]["required_slide_fields"]:
                if r.get(field) is None:
                    issue("Q01", "MISSING_REQUIRED_SLIDE_FIELD", field, sid)
            mapping = r.get("semantic_mapping", {}).get("mapping_status", "NOT_ASSESSED")
            if mapping != "SOURCE_BOUND":
                issue("Q03", "SEMANTIC_MAPPING_UNRESOLVED", mapping + "; retain original decision", sid, "NOT_ASSESSED")
            for prior in r.get("visual_QA", []):
                cid = prior.get("check_id")
                if cid in matrix and prior.get("result") in ("FAIL", "NOT_ASSESSED", "PENDING"):
                    issue(cid, "INHERITED_" + cid, prior.get("basis", "Existing QA concern"), sid, prior["result"])
            obs = obs_by_index.get(i, {})
            observed = fact(obs.get("evidence"))
            slide_sha = hashlib.sha256(archive.read(page["slide_part"])).hexdigest()
            observed_ok = (isinstance(observed, dict) and observed.get("slide_id") == sid and
                           observed.get("slide_index") == i and observed.get("slide_sha256") == slide_sha and
                           obs.get("preview_sha256") == previews_by_index.get(i, {}).get("sha256") and render_bound)
            family = observed.get("observed_family") if observed_ok else None
            if not isinstance(family, str) or not family.strip():
                issue("Q01", "ACTUAL_COMPOSITION_NOT_VERIFIED", "Declared/recommended family is not rendered evidence", sid)
                family = None
            c = r.get("composition_family") or {}
            norm = normalized_family(family)
            fallback = norm == "GENERIC_CARD_GRID" or norm in ("TEXT_ONLY", "TEXT_ONLY_FALLBACK") or c.get("generic_fallback") is True
            reason = c.get("fallback_reason")
            provenance = c.get("fallback_reason_provenance")
            fallback_status = "NOT_USED"
            if fallback:
                if not isinstance(reason, str) or not reason.strip() or not provenance or provenance == "NOT_APPLICABLE":
                    fallback_status = "SILENT_GENERIC_FALLBACK"
                    issue("Q06", fallback_status, "Missing fallback reason or source stage", sid)
                elif "RETROSPECTIVE" in provenance:
                    fallback_status = "RETROSPECTIVE_ONLY"
                    issue("Q06", "ORIGINAL_FALLBACK_RECEIPT_NOT_ASSESSED", provenance, sid, "NOT_ASSESSED")
                elif not c.get("candidate_evaluation"):
                    fallback_status = "EXCEPTION_EVIDENCE_INCOMPLETE"
                    issue("Q06", "FALLBACK_ALTERNATIVES_MISSING", "Rejected alternatives must remain visible", sid)
                else:
                    fallback_status = "EXPLICIT_JUSTIFIED"
            d = r.get("illustration_decision")
            decision = None
            if isinstance(d, dict) and type(d.get("illustration_needed")) is bool and d.get("decision_reason") and d.get("slide_id") == sid:
                source = d.get("selected_source")
                if not d["illustration_needed"] and source == "NONE" and not d.get("asset_ids") and d.get("illustration_role") == "NONE":
                    decision = "NO_ILLUSTRATION_NEEDED"
                elif d["illustration_needed"]:
                    decision = {"EXISTING_SCIENTIFIC_SOURCE_FIGURE": "SOURCE_FIGURE_REQUIRED",
                                "APPROVED_PROJECT_ASSET": "APPROVED_ASSET_SELECTED",
                                "APPROVED_DETERMINISTIC_LIBRARY": "APPROVED_ASSET_SELECTED",
                                "AI_CONCEPTUAL_CANDIDATE": "AI_CONCEPTUAL_CANDIDATE_POSSIBLE",
                                "NONE": "ILLUSTRATION_RECOMMENDED"}.get(source)
            if decision not in DECISIONS:
                issue("Q01", "MISSING_ILLUSTRATION_DECISION", "Explicit consistent boolean decision and rationale required", sid)
            d = d if isinstance(d, dict) else {}
            if (d.get("scientific_evidence") is not False and
                    d.get("selected_source") != "EXISTING_SCIENTIFIC_SOURCE_FIGURE"):
                issue("Q05", "ILLUSTRATION_EVIDENCE_FLAG_INVALID", "Only a registered source figure may be scientific evidence", sid)
            records.append({"slide_id": sid, "slide_index": i, "page_role": role if isinstance(role, str) and role in roles else None,
                            "composition_family": family, "semantic_mapping_status": mapping,
                            "generic_fallback_status": fallback_status, "fallback_used": fallback,
                            "fallback_reason": reason, "fallback_source_stage": provenance,
                            "illustration_decision": decision, "illustration_role": d.get("illustration_role"),
                            "illustration_source_type": d.get("selected_source"), "approved_asset_ids": [],
                            "visual_focal_point_status": "PENDING", "typography_hierarchy_status": "PENDING",
                            "art_direction_execution_status": "NOT_ASSESSED", "layout_repetition_family": norm,
                            "structure_pattern": structure[i-1], "scientific_qa_status": scientific,
                            "geometry_qa_status": geometry, "typography_minimum_status": typography,
                            "warnings": [], "human_review_reasons": [],
                            "illustration_subject_summary": d.get("illustration_subject"),
                            "allowed_visual_types": safety["allowed_subject_classes"],
                            "prohibited_visual_types": safety["forbidden_evidence_types"],
                            "source_priority": authority["illustration_decision.schema.json"]["properties"]["source_priority"]["const"],
                            "human_approval_required": True, "insertion_allowed_now": False,
                            "design_record": r, "observed_slide_sha256": slide_sha})

    # Check every actual media relationship, including inherited layout/master media.
    registry = manifest.get("assets", [])
    used_ids = defaultdict(set)
    actual_media_slides = set()
    for relation in package["asset_relationships"]:
        if relation["category"] != "media":
            continue
        chains = [c for c in package["slide_relationship_chains"] if relation["source_part"] in
                  (c["slide_part"], c["layout_part"], c["master_part"], c["theme_part"])]
        if not chains:
            issue("Q04", "UNAUTHORIZED_PRESENTATION_ASSET", "Unscoped media relationship: " + relation["source_part"])
        for chain in chains:
            rec = records[chain["slide_index"]-1]
            sid = rec["slide_id"]
            actual_media_slides.add(sid)
            matches = [a for a in registry if a.get("slide_id") == sid and
                       any(m.get("part") == relation["target_part"] and m.get("sha256") == relation.get("target_sha256")
                           for m in a.get("media", []))]
            allowed = False
            for a in matches:
                grant = fact(a.get("authorization"))
                placement = fact(a.get("placement"))
                source_type = a.get("source_type")
                if not isinstance(grant, dict) or not isinstance(placement, dict):
                    continue
                aid = a.get("asset_id")
                if (placement.get("asset_id") != aid or placement.get("slide_id") != sid or
                        not any(m.get("part") == relation["target_part"] and m.get("sha256") == relation.get("target_sha256")
                                for m in placement.get("media", []))):
                    continue
                evidence_flag = grant.get("scientific_evidence")
                if source_type == "EXISTING_SCIENTIFIC_SOURCE_FIGURE":
                    ok = (evidence_flag is True and grant.get("source_id") and grant.get("figure_id") and
                          grant.get("scientific_qa_status") == "PASS" and grant.get("deck_sha256") == deck_sha and
                          sid in grant.get("slide_ids", []) and verify(a.get("source_asset")) and
                          a["source_asset"]["sha256"] == grant.get("source_sha256"))
                elif source_type in ("APPROVED_PROJECT_ASSET", "APPROVED_DETERMINISTIC_LIBRARY", "AI_CONCEPTUAL_CANDIDATE"):
                    ok = evidence_flag is False and grant.get("presentation_only") is True
                    events = fact(a.get("approval_events"))
                    events = [e for e in events if e.get("asset_id") == aid] if isinstance(events, list) else []
                    ok = (ok and grant.get("status") == "APPROVED" and grant.get("asset_id") == aid and
                          grant.get("project_id") == manifest.get("project_id") and
                          grant.get("deck_content_sha256") == fact(manifest.get("deck_content_binding")) and
                          bool(grant.get("deck_content_sha256")) and sid in grant.get("slide_ids", []) and
                          grant.get("asset_sha256") == a.get("source_sha256") and
                          grant.get("placement_role") == rec["illustration_role"] and
                          a.get("transform") in grant.get("allowed_transforms", []) and
                          grant.get("reviewed_preview_sha256") == previews_by_index.get(rec["slide_index"], {}).get("sha256") and
                          grant.get("approved_by") and grant.get("approved_at") and
                          grant.get("decision_evidence_ref") and
                          bool(events) and events[-1] == grant and
                          verify(a.get("source_asset")) and
                          a["source_asset"]["sha256"] == grant.get("asset_sha256") and
                          placement.get("source_sha256") == grant.get("asset_sha256"))
                    if source_type == "AI_CONCEPTUAL_CANDIDATE":
                        candidate = fact(a.get("candidate"))
                        safety_review = fact(a.get("safety_review"))
                        ok = (ok and isinstance(candidate, dict) and candidate.get("approval_status") == "CANDIDATE" and
                              candidate.get("scientific_evidence") is False and candidate.get("presentation_only") is True and
                              candidate.get("asset_id") == aid and candidate.get("asset_sha256") == grant.get("asset_sha256"))
                        safe = (isinstance(safety_review, dict) and safety_review.get("status") == "PASS" and
                                safety_review.get("asset_sha256") == grant.get("asset_sha256") and
                                safety_review.get("forbidden_evidence_appearance") is False and
                                safety_review.get("reviewer") and safety_review.get("reviewed_at"))
                        if not safe:
                            issue("Q05", "GENERATED_ASSET_SAFETY_NOT_VERIFIED", "Exact generated image needs visual/scientific safety review", sid)
                            ok = False
                    if source_type == "AI_CONCEPTUAL_CANDIDATE" and (evidence_flag is not False or
                            grant.get("visual_type") in safety["forbidden_evidence_types"] or
                            rec["illustration_decision"] == "SOURCE_FIGURE_REQUIRED"):
                        issue("Q05", "UNSAFE_SCIENTIFIC_LOOKING_ILLUSTRATION", "Generated asset violates evidence boundary", sid)
                        ok = False
                else:
                    ok = False
                if ok:
                    allowed = True
                    used_ids[sid].add(aid)
                if evidence_flag is not False and source_type != "EXISTING_SCIENTIFIC_SOURCE_FIGURE":
                    issue("Q05", "UNSAFE_SCIENTIFIC_LOOKING_ILLUSTRATION", "Decorative asset must preserve scientific_evidence=false", sid)
            if not allowed:
                issue("Q04", "UNAUTHORIZED_PRESENTATION_ASSET", relation["target_part"], sid)
    for rec in records:
        sid = rec["slide_id"]
        rec["approved_asset_ids"] = sorted(used_ids[sid])
        expected = set((rec["design_record"].get("illustration_decision") or {}).get("asset_ids", []))
        if expected != used_ids[sid]:
            issue("Q04", "ASSET_DECISION_REALIZATION_MISMATCH", "Selected IDs differ from authorized media on actual page", sid)
        if rec["illustration_decision"] == "SOURCE_FIGURE_REQUIRED" and not used_ids[sid]:
            issue("Q04", "REQUIRED_SOURCE_FIGURE_MISSING", "AI artwork cannot fill the evidence requirement", sid)

    # Presence, consumption and rendered observation are intentionally independent.
    art_rows = []
    art = manifest.get("art_direction", {})
    brief_valid = verify(art.get("brief"))
    brief = {}
    if brief_valid:
        import yaml
        brief = yaml.safe_load(Path(art["brief"]["path"]).read_text(encoding="utf-8-sig"))
        brief_valid = isinstance(brief, dict) and brief.get("approval", {}).get("status") == "approved"
    direction = brief.get("direction", {}) if brief_valid else {}
    for field in ART_FIELDS:
        entries = [x for x in art.get("fields", []) if x.get("field") == field]
        for rec in records:
            candidates = [x for x in entries if rec["slide_id"] in x.get("slide_ids", [])]
            row = candidates[0] if len(candidates) == 1 else {}
            declared = (direction.get("primary_palette", []) + direction.get("secondary_palette", []) if field == "palette"
                        else brief.get("approved_asset_ids") if field == "approved_assets" else direction.get(field))
            consumer = fact(row.get("consumer_evidence"))
            observation = fact(row.get("render_evidence"))
            consumed = bool(row.get("consumer") and consumer is not None and consumer == row.get("consumer_value"))
            observed = bool(observation is not None and observation == row.get("render_value") and
                            row.get("render_deck_sha256") == deck_sha)
            na_evidence = fact(row.get("not_applicable_evidence"))
            na = (row.get("not_applicable") is True and bool(row.get("not_applicable_reason")) and
                  na_evidence is not None and na_evidence == row.get("not_applicable_value"))
            art_rows.append({"slide_id": rec["slide_id"], "field": field, "declared": declared,
                             "consumed": consumed, "observed_in_render": observed, "not_applicable": na,
                             "not_applicable_reason": row.get("not_applicable_reason"),
                             "consumer": row.get("consumer"), "consumer_evidence": row.get("consumer_evidence"),
                             "render_evidence": row.get("render_evidence"), "limitation": row.get("limitation"),
                             "execution_status": "NOT_APPLICABLE" if na else
                                 "OBSERVED" if consumed and observed else "NOT_ASSESSED"})
    for rec in records:
        rows = [a for a in art_rows if a["slide_id"] == rec["slide_id"]]
        incomplete = [a["field"] for a in rows if a["execution_status"] == "NOT_ASSESSED"]
        rec["art_direction_execution_status"] = "PARTIAL" if incomplete else "RECORDED_AND_OBSERVED"
        if incomplete:
            issue("Q18", "ART_DIRECTION_NOT_EXECUTED_OR_UNVERIFIED", ", ".join(incomplete), rec["slide_id"], "NOT_ASSESSED")
    families = [r["layout_repetition_family"] for r in records]
    counts = Counter(families)
    generic = counts["GENERIC_CARD_GRID"]
    illustrated = len(actual_media_slides)
    macro_run = longest_run(families)
    generic_run = longest_run(families, "GENERIC_CARD_GRID")
    if macro_run > 3 or max(counts.values(), default=0) / (count or 1) > .35:
        issue("Q09", "MACRO_COMPOSITION_REPETITION", "Advisory run >3 or dominant share >35%; no aesthetic block")
    if generic / (count or 1) > .25:
        issue("Q10", "HIGH_GENERIC_CARD_DEPENDENCE", "Advisory generic-card ratio >25%")
    # Three consecutive cards trigger an advisory, not an aesthetic verdict.
    if generic_run >= 3:
        issue("Q10", "CONSECUTIVE_GENERIC_CARDS", f"{generic_run} pages; Early warning, agreed threshold NOT_ASSESSED")
    patterns = defaultdict(list)
    repeated_families = defaultdict(list)
    for rec in records:
        patterns[rec["structure_pattern"]].append(rec["slide_index"])
        repeated_families[rec["layout_repetition_family"]].append(rec["slide_index"])
    nonconsecutive = {k: v for k, v in repeated_families.items() if any(b-a > 1 for a, b in zip(v, v[1:]))}
    repeated_patterns = {k: v for k, v in patterns.items() if len(v) > 1}
    if nonconsecutive or repeated_patterns:
        issue("Q09", "REPEATED_COMPOSITION_PATTERN", "Non-consecutive family or label-independent macro geometry repetition")
    if longest_run([r["structure_pattern"] for r in records]) >= 3:
        issue("Q09", "CONSECUTIVE_NEAR_IDENTICAL_STRUCTURE", "At least three pages within 2.5% macro-geometry tolerance")
    if illustrated / (count or 1) > .4:
        issue("Q15", "ILLUSTRATION_DENSITY_HIGH", "Proposed >40%; appropriateness still needs review")

    # Signed approval is specific to this exact deck and render set, never borrowed.
    render_hash = digest({"previews": [p.get("sha256") for p in previews], "contact_sheet": sheet.get("sha256")})
    for cid, check in matrix.items():
        if check["severity"] != "HUMAN_REVIEW":
            continue
        review = fact(manifest.get("human_reviews", {}).get(cid))
        approved = (isinstance(review, dict) and review.get("status") == "APPROVED" and
                    review.get("reviewer") and review.get("reviewed_at") and
                    review.get("deck_sha256") == deck_sha and review.get("render_set_sha256") == render_hash and
                    review.get("slide_ids") == [r["slide_id"] for r in records] and
                    (cid != "Q20" or review.get("scientific_approval") is True and review.get("visual_approval") is True))
        if not approved:
            issue(cid, check["dimension"].upper(), check["method"], result="PENDING")
        for rec in records:
            if cid == "Q11":
                rec["visual_focal_point_status"] = "PASS" if approved else "PENDING"
            elif cid == "Q13":
                rec["typography_hierarchy_status"] = "PASS" if approved else "PENDING"
    for path, expected in checked.items():
        if sha256(path) != expected:
            raise ValueError("INPUT_CHANGED_DURING_READ_ONLY_EVALUATION: " + path)
    checks = []
    for cid, check in matrix.items():
        active = [f for f in findings if f["check_id"] == cid]
        result = "FAIL" if any(f["result"] == "FAIL" for f in active) else (
            "NOT_ASSESSED" if any(f["result"] == "NOT_ASSESSED" for f in active) else "PENDING" if active else "PASS")
        checks.append({**check, "result": result, "finding_count": len(active)})
    blocked = any(f["severity"] == "BLOCK" for f in findings)
    human = any(f["severity"] == "HUMAN_REVIEW" for f in findings) or any(
        f["code"] in ("SEMANTIC_MAPPING_UNRESOLVED", "ORIGINAL_FALLBACK_RECEIPT_NOT_ASSESSED") for f in findings)
    status = ("VISUAL_QUALITY_BLOCKED" if blocked else "VISUAL_QUALITY_HUMAN_REVIEW_REQUIRED" if human else
              "VISUAL_QUALITY_PASS_WITH_WARNINGS" if findings else "VISUAL_QUALITY_PASS")
    for rec in records:
        relevant = [f for f in findings if f["slide_id"] in (None, rec["slide_id"])]
        rec["warnings"] = [f["code"] for f in relevant if f["severity"] == "WARN"]
        rec["blocking_reasons"] = [f["code"] for f in relevant if f["severity"] == "BLOCK"]
        rec["human_review_reasons"] = [f["code"] for f in relevant if f["severity"] == "HUMAN_REVIEW" or f["result"] == "NOT_ASSESSED"]
        if rec["fallback_used"]:
            rec["human_review_reasons"].append("REVIEW_GENERIC_FALLBACK_EXCEPTION")
    art_coverage = {field: {"record_count": count, "declared_count": sum(a["declared"] is not None for a in art_rows if a["field"] == field),
                           **{key: sum(a[key] for a in art_rows if a["field"] == field)
                              for key in ("consumed", "observed_in_render", "not_applicable")}}
                    for field in ART_FIELDS}
    summary = {"slide_count": count, "page_role_distribution": dict(Counter(str(r["page_role"]) for r in records)),
               "composition_family_distribution": dict(counts), "generic_fallback_count": sum(r["fallback_used"] for r in records),
               "generic_card_count": generic, "generic_card_ratio": generic / (count or 1),
               "max_consecutive_generic_cards": generic_run, "max_consecutive_composition_family": macro_run,
               "illustration_page_count": illustrated, "illustration_ratio": illustrated / (count or 1),
               "page_role_coverage": sum(isinstance(r["page_role"], str) and r["page_role"] in roles for r in records) / (count or 1),
               "illustration_decision_coverage": sum(r["illustration_decision"] in DECISIONS for r in records) / (count or 1),
               "art_direction_execution_coverage": art_coverage, "human_review_required": human or blocked,
               "non_consecutive_repeated_families": nonconsecutive, "repeated_structure_patterns": repeated_patterns,
               "same_structure_different_labels": {k: v for k, v in repeated_patterns.items()
                    if len({families[i-1] for i in v}) > 1},
               "per_role_composition": {role: dict(Counter(r["layout_repetition_family"] for r in records if r["page_role"] == role))
                                        for role in sorted(roles)},
               "agreed_threshold_evaluation": "NOT_ASSESSED", "threshold_status": "PROPOSED_NOT_AGREED"}
    return {"schema_version": VERSION, "mode": mode.upper(), "experimental": False, "default_off": True,
            "canonical": True, "production_active": mode == "enforce", "ai_generation_enabled": False,
            "status": status, "deck_sha256": deck_sha, "render_set_sha256": render_hash,
            "deck_floor_acceptance": "NOT_CERTIFIED" if blocked or human else "PROCESS_FLOOR_SATISFIED",
            "route": manifest.get("route", "local_quality_evaluation"),
            "legacy_limitations": manifest.get("legacy_limitations", []),
            "other_routes": "FLOOR_NOT_ASSESSED_FOR_ROUTE", "cross_project_validation": "SECOND_INDEPENDENT_REAL_PROJECT_REQUIRED",
            "authority": authority["lock"], "checks": checks, "findings": findings,
            "summary": summary, "slide_records": records, "art_direction_records": art_rows,
            "verified_input_sha256": checked, "deck_modified": False, "new_render_performed": False,
            "illustrations_generated": 0, "requests_dispatched": 0}


def write_outputs(report, output):
    output = Path(output).resolve()
    if ROOT in output.parents and not any((ROOT / name).resolve() in output.parents for name in ("staging", "output")):
        raise ValueError("Use an external project output or a new staging directory")
    output.mkdir(parents=True, exist_ok=False)
    for name, value in {
        "visual_quality_report.json": {k: v for k, v in report.items() if k not in ("slide_records", "art_direction_records")},
        "slide_quality_records.json": report["slide_records"], "deck_quality_summary.json": report["summary"],
        "art_direction_execution_report.json": {"records": report["art_direction_records"], "coverage": report["summary"]["art_direction_execution_coverage"]},
        "illustration_decision_report.json": {"coverage": report["summary"]["illustration_decision_coverage"],
            "generation_enabled": False, "decisions": [{k: r[k] for k in ("slide_id", "slide_index", "illustration_decision",
                "illustration_role", "illustration_source_type", "illustration_subject_summary", "approved_asset_ids",
                "allowed_visual_types", "prohibited_visual_types", "source_priority", "human_approval_required", "insertion_allowed_now")}
                for r in report["slide_records"]]},
    }.items():
        (output / name).write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf8")
    lines = ["# Visual quality process floor", "", report["status"], "", "Deck acceptance: " + report["deck_floor_acceptance"],
             "", "Read-only evidence evaluation; no scientific re-adjudication, deck edits or asset generation.",
             "", "| Check | Severity | Result |", "|---|---|---|"]
    lines += [f"| {c['check_id']} {c['dimension']} | {c['severity']} | {c['result']} |" for c in report["checks"]]
    lines += ["", "## Findings", ""] + [f"- {f['severity']} / {f['slide_id'] or 'DECK'} / {f['code']}: {f['basis']}" for f in report["findings"]]
    lines += ["", "Cross-project visual generalization requires a second independent real academic deck."]
    (output / "visual_quality_report.md").write_text("\n".join(lines) + "\n", encoding="utf8")
    with (output / "human_review_queue.csv").open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=["slide_id", "check_id", "severity", "result", "code", "basis", "deck_sha256", "render_set_sha256"])
        writer.writeheader()
        for f in report["findings"]:
            writer.writerow({**f, "deck_sha256": report["deck_sha256"], "render_set_sha256": report["render_set_sha256"]})
    (output / "quality_floor_execution_log.md").write_text(
        "# Execution log\n\nRead-only evaluation. Inputs verified before and after evaluation.\n\n"
        f"Deck SHA-256: {report['deck_sha256']}\n\nVerified files: {len(report['verified_input_sha256'])}\n\n"
        "No deck mutation, renderer call, network dispatch or asset generation.\n\n"
        "Other production routes: FLOOR_NOT_ASSESSED_FOR_ROUTE.\n", encoding="utf8")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--mode", choices=("off", "shadow", "enforce"), default="off")
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.mode == "off":
        print("DEFAULT_OFF: evaluator not invoked; no input/output access")
        return
    if not args.manifest or not args.output:
        parser.error("--manifest and --output required in shadow mode")
    report = evaluate(read_json(args.manifest), mode=args.mode)
    write_outputs(report, args.output)
    print(report["status"])


if __name__ == "__main__":
    main()
