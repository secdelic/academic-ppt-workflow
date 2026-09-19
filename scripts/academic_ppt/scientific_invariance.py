"""Preserved source/claim binding checks shared by native composition and assembly."""
import re

def normalized(text: str) -> str:
    return " ".join(text.split())

def source_binding_issues(slides: list[dict], originals: list[dict], claims: list[dict], manifest: list[dict], notes: dict) -> list[str]:
    """Verify preserved existing claims, which may be full excerpts, not titles.

    No new claim is authorized by this check. All inherited scientific fields
    must match the source baseline exactly and every registered excerpt must occur in the source baseline notes.
    """
    issues=[]
    old={s["slide_id"]:s for s in originals}
    registry={c["claim_id"]:c for c in claims}
    sources={s["source_id"]:s for s in manifest}
    fields=("slide_title","single_key_message","source_ids","claim_ids","short_source_label",
            "speaker_notes","prohibited_overstatement","figure_ids","chart_data")
    for slide in slides:
        sid=slide["slide_id"]
        if sid not in old or any(slide.get(k)!=old[sid].get(k) for k in fields):
            issues.append(sid+": inherited scientific field changed")
            continue
        if not slide["claim_ids"] or not set(slide["source_ids"])<=sources.keys():
            issues.append(sid+": missing source/claim identity")
        for cid in slide["claim_ids"]:
            claim=registry.get(cid)
            if not claim or claim["source_id"] not in slide["source_ids"]:
                issues.append(sid+": unknown or mismatched claim")
            elif (claim.get("source_sha256")!=sources[claim["source_id"]]["sha256"]
                  or not claim.get("claim_text") or normalized(claim["claim_text"]) not in normalized(notes[sid])):
                issues.append(sid+": frozen excerpt/source hash mismatch")
    return issues
