from __future__ import annotations

import hashlib
import hmac
import json
import os
import re
import secrets
import shutil
import tempfile
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Mapping

from .utils import sha256_file


DELTA_STATUSES = frozenset({"UNCHANGED", "MODIFIED", "NEW", "REMOVED"})
PROJECT_STATE_SCHEMA_VERSION = "academic-ppt-project-state-payload/2"
REQUIRED_PROJECT_STATE_FIELDS = frozenset(
    {
        "cache_contract_fingerprints",
        "source_manifest",
        "source_hashes",
        "parsed_document_objects",
        "evidence_registry",
        "slide_specs",
        "visual_specs",
        "master_layout_map",
        "font_map",
        "image_registry",
        "source_bindings",
        "previous_qa_status",
    }
)
REQUIRED_CACHE_CONTRACT_FINGERPRINTS = frozenset(
    {
        "brief_fingerprint",
        "scientific_definition_fingerprint",
        "extractor_fingerprint",
        "config_fingerprint",
        "code_fingerprint",
    }
)
_BLOCKED_CACHE_PARTS = frozenset(
    {"tests", "regression", "benchmark", "output", "archive"}
)
_PROJECT_KEY_RE = re.compile(r"^PRJ-[A-F0-9]{32}$")
_GENERATION_ID_RE = re.compile(r"^GEN-[a-f0-9]{16,64}$")
_SHA256_RE = re.compile(r"^[a-fA-F0-9]{64}$")


class CachePolicyError(ValueError):
    """Raised when a cache location or identifier violates local-cache policy."""


class CacheStateError(RuntimeError):
    """Raised when a cache generation is missing, malformed, or cannot promote."""


@dataclass(frozen=True)
class SourceDelta:
    source_key: str
    status: str
    previous_sha256: str
    current_sha256: str

    def to_dict(self) -> dict[str, str]:
        return {
            "source_key": self.source_key,
            "status": self.status,
            "previous_sha256": self.previous_sha256,
            "current_sha256": self.current_sha256,
        }


def _utc_timestamp() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _is_unc_path(path: Path | str) -> bool:
    raw = os.fspath(path).replace("/", "\\")
    return raw.startswith("\\\\")


def _is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _assert_json_value(value: Any, *, label: str = "cache state") -> None:
    """Reject non-JSON objects explicitly instead of relying on implicit coercion."""

    try:
        json.dumps(value, ensure_ascii=False, allow_nan=False)
    except (TypeError, ValueError) as exc:
        raise CacheStateError(f"{label} must contain JSON values only") from exc


def _canonical_json_sha256(value: Any) -> str:
    _assert_json_value(value, label="fingerprint input")
    payload = json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def _file_set_fingerprint(root: Path, files: Iterable[Path]) -> str:
    rows: list[dict[str, str]] = []
    resolved_root = root.resolve()
    for raw_path in sorted({Path(item).resolve() for item in files}, key=str):
        if not raw_path.is_file():
            continue
        try:
            relative = raw_path.relative_to(resolved_root).as_posix()
        except ValueError as exc:
            raise CachePolicyError(
                f"Fingerprint file is outside the repository: {raw_path}"
            ) from exc
        rows.append({"path": relative, "sha256": sha256_file(raw_path).lower()})
    if not rows:
        raise CacheStateError("Fingerprint file set is empty")
    return _canonical_json_sha256(rows)


_SCIENTIFIC_DEFINITION_KEYS = frozenset(
    {
        "analysis_plan",
        "baseline",
        "baseline_definition",
        "cohort",
        "cohort_definition",
        "eligibility",
        "estimand",
        "exposure",
        "exposure_definition",
        "intervention",
        "intervention_definition",
        "missing_data_strategy",
        "missingness_strategy",
        "model",
        "narrative_mode",
        "outcome",
        "outcome_definition",
        "primary_outcome",
        "secondary_outcomes",
        "statistical_model",
        "time_window",
        "validation_strategy",
    }
)


def _scientific_definition_contract(value: Any) -> Any:
    """Return only explicitly scientific brief fields, preserving their paths."""

    if isinstance(value, Mapping):
        selected: dict[str, Any] = {}
        for raw_key, raw_value in sorted(value.items(), key=lambda item: str(item[0])):
            key = str(raw_key)
            child = _scientific_definition_contract(raw_value)
            if key.casefold() in _SCIENTIFIC_DEFINITION_KEYS:
                selected[key] = raw_value
            elif child not in ({}, []):
                selected[key] = child
        return selected
    if isinstance(value, list):
        children = [_scientific_definition_contract(item) for item in value]
        return [item for item in children if item not in ({}, [])]
    return {}


def build_cache_contract_fingerprints(
    *,
    repo_root: Path,
    brief: Mapping[str, Any],
    config_root: Path | None = None,
) -> dict[str, str]:
    """Bind cache reuse to the exact brief, science, extractor, config, and code.

    Only digests are persisted, so a private brief is never copied into cache
    metadata.  The full brief digest deliberately makes the fast path
    conservative: any brief change requires a fresh ``full_validation``.
    """

    repo = repo_root.resolve()
    config_dir = (config_root or (repo / "config")).resolve()
    config_files = [path for path in config_dir.rglob("*") if path.is_file()]
    module_dir = repo / "scripts" / "academic_ppt"
    code_files = [path for path in module_dir.rglob("*.py") if path.is_file()]
    code_files.extend(
        path
        for path in (repo / "scripts").glob("*.ps1")
        if path.is_file()
    )
    extractor = module_dir / "extractors.py"
    if not extractor.is_file():
        raise CacheStateError("Canonical extractor module is missing")
    # ``fast_enhance`` is an execution request, not a scientific definition.
    # Excluding it lets one validated baseline accept a succession of bounded
    # change requests while every scientific brief field remains frozen.
    stable_brief = {
        str(key): value for key, value in brief.items()
        if str(key) not in {"fast_enhance"}
    }
    result = {
        "brief_fingerprint": _canonical_json_sha256(stable_brief),
        "scientific_definition_fingerprint": _canonical_json_sha256(
            _scientific_definition_contract(brief)
        ),
        "extractor_fingerprint": sha256_file(extractor).lower(),
        "config_fingerprint": _file_set_fingerprint(repo, config_files),
        "code_fingerprint": _file_set_fingerprint(repo, code_files),
    }
    validate_cache_contract_fingerprints(result)
    return result


def validate_cache_contract_fingerprints(value: Any) -> None:
    if not isinstance(value, Mapping):
        raise CacheStateError("cache_contract_fingerprints must be an object")
    missing = sorted(REQUIRED_CACHE_CONTRACT_FINGERPRINTS.difference(value))
    unexpected = sorted(set(value).difference(REQUIRED_CACHE_CONTRACT_FINGERPRINTS))
    if missing or unexpected:
        details = []
        if missing:
            details.append("missing=" + ",".join(missing))
        if unexpected:
            details.append("unexpected=" + ",".join(unexpected))
        raise CacheStateError(
            "cache_contract_fingerprints fields are invalid: " + "; ".join(details)
        )
    invalid = sorted(
        key for key, digest in value.items() if not _SHA256_RE.fullmatch(str(digest))
    )
    if invalid:
        raise CacheStateError(
            "cache contract fingerprint is not SHA-256: " + ", ".join(invalid)
        )


def _atomic_write_json(path: Path, value: Any) -> None:
    _assert_json_value(value, label=str(path.name))
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        with temporary.open("x", encoding="utf-8", newline="\n") as handle:
            json.dump(value, handle, ensure_ascii=False, indent=2, allow_nan=False)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        if temporary.exists():
            temporary.unlink()


def validate_cache_root(repo_root: Path, cache_root: Path | None = None) -> Path:
    """Return a safe repository-local cache root.

    Clinical cache state is deliberately restricted to ``<repo>/.cache``.  This
    also fails closed for common publication/test/output trees and UNC paths.
    ``resolve`` is intentional: an existing symlink under ``.cache`` cannot be
    used to redirect protected state outside the repository.
    """

    if _is_unc_path(repo_root):
        raise CachePolicyError("UNC repositories cannot host clinical project cache")
    repo = repo_root.resolve()
    candidate_raw = cache_root or (repo / ".cache" / "project_state")
    if _is_unc_path(candidate_raw):
        raise CachePolicyError("UNC cache paths are prohibited")
    candidate = Path(candidate_raw).resolve()
    allowed_parent = (repo / ".cache").resolve()
    if not _is_relative_to(candidate, allowed_parent):
        raise CachePolicyError("Project cache must be inside the repository .cache tree")
    relative_parts = {part.casefold() for part in candidate.relative_to(repo).parts}
    blocked = sorted(relative_parts.intersection(_BLOCKED_CACHE_PARTS))
    if blocked:
        raise CachePolicyError(
            "Project cache path contains prohibited tree component(s): "
            + ", ".join(blocked)
        )
    return candidate


def _load_or_create_key_secret(cache_root: Path) -> bytes:
    secret_path = cache_root / ".project_key_secret.json"
    cache_root.mkdir(parents=True, exist_ok=True)
    if secret_path.is_file():
        try:
            document = json.loads(secret_path.read_text(encoding="utf-8"))
            secret_hex = document["secret_hex"]
            secret = bytes.fromhex(secret_hex)
        except (KeyError, TypeError, ValueError, json.JSONDecodeError) as exc:
            raise CacheStateError("Project-key secret is malformed") from exc
        if len(secret) < 32:
            raise CacheStateError("Project-key secret is too short")
        return secret

    secret = secrets.token_bytes(32)
    payload = {
        "schema_version": "academic-ppt-project-key-secret/1",
        "secret_hex": secret.hex(),
    }
    try:
        # Atomic replace handles concurrent first writers.  A later writer will
        # reread the final value so all processes derive the same project key.
        _atomic_write_json(secret_path, payload)
    except OSError:
        if not secret_path.is_file():
            raise
    final = json.loads(secret_path.read_text(encoding="utf-8"))
    try:
        return bytes.fromhex(final["secret_hex"])
    except (KeyError, TypeError, ValueError) as exc:
        raise CacheStateError("Project-key secret is malformed") from exc


def derive_opaque_project_key(project_identity: str, *, secret: bytes | str) -> str:
    """Derive a stable opaque cache key without storing the project identity."""

    identity = str(project_identity).strip()
    if not identity:
        raise CachePolicyError("Project identity is required")
    secret_bytes = secret.encode("utf-8") if isinstance(secret, str) else bytes(secret)
    if len(secret_bytes) < 16:
        raise CachePolicyError("Opaque project keys require at least 16 secret bytes")
    digest = hmac.new(
        secret_bytes,
        identity.encode("utf-8"),
        digestmod=hashlib.sha256,
    ).hexdigest()[:32].upper()
    return f"PRJ-{digest}"


def build_source_manifest(
    sources: Mapping[str, Path],
    *,
    relative_to: Path | None = None,
) -> dict[str, Any]:
    """Hash source files into a deterministic, JSON-only manifest.

    ``source_key`` is the comparison identity.  Callers should use registered
    source IDs rather than patient-facing file names.  Absolute paths are never
    persisted; ``relative_path`` is present only when ``relative_to`` is given.
    """

    base = relative_to.resolve() if relative_to is not None else None
    rows: list[dict[str, Any]] = []
    for source_key, raw_path in sorted(sources.items(), key=lambda item: item[0]):
        key = str(source_key).strip()
        if not key:
            raise CacheStateError("Source manifest contains an empty source_key")
        path = Path(raw_path).resolve()
        if not path.is_file():
            raise FileNotFoundError(f"Source file does not exist: {path}")
        relative_path = ""
        if base is not None:
            if not _is_relative_to(path, base):
                raise CachePolicyError(f"Source is outside the declared input root: {path}")
            relative_path = path.relative_to(base).as_posix()
        rows.append(
            {
                "source_key": key,
                "relative_path": relative_path,
                "sha256": sha256_file(path).lower(),
                "size_bytes": path.stat().st_size,
            }
        )
    return {
        "schema_version": "academic-ppt-source-manifest/1",
        "hash_algorithm": "sha256",
        "sources": rows,
    }


def _manifest_index(manifest: Mapping[str, Any] | None) -> dict[str, str]:
    if manifest is None:
        return {}
    if str(manifest.get("hash_algorithm", "sha256")).casefold() != "sha256":
        raise CacheStateError("Only SHA-256 source manifests are supported")
    rows = manifest.get("sources", [])
    if isinstance(rows, Mapping):
        iterable: Iterable[Any] = [
            {"source_key": key, **(value if isinstance(value, Mapping) else {})}
            for key, value in rows.items()
        ]
    elif isinstance(rows, list):
        iterable = rows
    else:
        raise CacheStateError("Source manifest 'sources' must be a list or object")
    result: dict[str, str] = {}
    for row in iterable:
        if not isinstance(row, Mapping):
            raise CacheStateError("Source manifest entries must be objects")
        key = str(row.get("source_key", "")).strip()
        digest = str(row.get("sha256", "")).strip().lower()
        if not key:
            raise CacheStateError("Source manifest contains an empty source_key")
        if key in result:
            raise CacheStateError(f"Duplicate source_key in manifest: {key}")
        if not _SHA256_RE.fullmatch(digest):
            raise CacheStateError(f"Invalid SHA-256 for source_key {key}")
        result[key] = digest
    return result


def compute_delta_manifest(
    previous_manifest: Mapping[str, Any] | None,
    current_manifest: Mapping[str, Any] | None,
) -> dict[str, Any]:
    """Classify source deltas using SHA-256 as the sole change authority."""

    previous = _manifest_index(previous_manifest)
    current = _manifest_index(current_manifest)
    deltas: list[SourceDelta] = []
    for source_key in sorted(set(previous).union(current)):
        before = previous.get(source_key, "")
        after = current.get(source_key, "")
        if not before:
            status = "NEW"
        elif not after:
            status = "REMOVED"
        elif before == after:
            status = "UNCHANGED"
        else:
            status = "MODIFIED"
        deltas.append(SourceDelta(source_key, status, before, after))
    counts = {
        status: sum(delta.status == status for delta in deltas)
        for status in sorted(DELTA_STATUSES)
    }
    return {
        "schema_version": "academic-ppt-delta-manifest/1",
        "hash_algorithm": "sha256",
        "entries": [delta.to_dict() for delta in deltas],
        "counts": counts,
        "parse_source_keys": [
            delta.source_key
            for delta in deltas
            if delta.status in {"NEW", "MODIFIED"}
        ],
        "unchanged_source_keys": [
            delta.source_key for delta in deltas if delta.status == "UNCHANGED"
        ],
        "removed_source_keys": [
            delta.source_key for delta in deltas if delta.status == "REMOVED"
        ],
    }


def build_project_state(
    *,
    cache_contract_fingerprints: Mapping[str, str],
    source_manifest: Mapping[str, Any],
    parsed_document_objects: Mapping[str, Any],
    evidence_registry: Mapping[str, Any] | list[Any],
    slide_specs: list[Mapping[str, Any]],
    visual_specs: list[Mapping[str, Any]],
    master_layout_map: Mapping[str, Any],
    font_map: Mapping[str, Any],
    image_registry: Mapping[str, Any] | list[Any],
    source_bindings: Mapping[str, Any] | list[Any],
    previous_qa_status: Mapping[str, Any],
    extra: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Build the complete v2.5.1 project-state cache contract.

    Source hashes are copied from the authoritative SHA-256 manifest so cache
    consumers never create a second change-detection authority.
    """

    source_hashes = _manifest_index(source_manifest)
    state: dict[str, Any] = {
        "schema_version": PROJECT_STATE_SCHEMA_VERSION,
        "cache_contract_fingerprints": dict(cache_contract_fingerprints),
        "source_manifest": dict(source_manifest),
        "source_hashes": source_hashes,
        "parsed_document_objects": dict(parsed_document_objects),
        "evidence_registry": evidence_registry,
        "slide_specs": list(slide_specs),
        "visual_specs": list(visual_specs),
        "master_layout_map": dict(master_layout_map),
        "font_map": dict(font_map),
        "image_registry": image_registry,
        "source_bindings": source_bindings,
        "previous_qa_status": dict(previous_qa_status),
    }
    if extra:
        collision = set(extra).intersection(state)
        if collision:
            raise CacheStateError(
                "Project-state extra fields cannot replace canonical fields: "
                + ", ".join(sorted(collision))
            )
        state.update(extra)
    validate_project_state(state)
    return state


def validate_project_state(state: Mapping[str, Any]) -> None:
    if not isinstance(state, Mapping):
        raise CacheStateError("Project state must be a JSON object")
    missing = sorted(REQUIRED_PROJECT_STATE_FIELDS.difference(state))
    if missing:
        raise CacheStateError(
            "Project state is missing required cache field(s): " + ", ".join(missing)
        )
    if state.get("schema_version") != PROJECT_STATE_SCHEMA_VERSION:
        raise CacheStateError(
            "Project state schema is incompatible; rebuild with full_validation"
        )
    _assert_json_value(state)
    validate_cache_contract_fingerprints(state.get("cache_contract_fingerprints"))
    manifest_hashes = _manifest_index(state.get("source_manifest"))
    cached_hashes = state.get("source_hashes")
    if not isinstance(cached_hashes, Mapping):
        raise CacheStateError("Project state source_hashes must be an object")
    normalised_hashes = {str(key): str(value).lower() for key, value in cached_hashes.items()}
    if normalised_hashes != manifest_hashes:
        raise CacheStateError("source_hashes must exactly match source_manifest SHA-256 values")


class ProjectCache:
    """Repository-local immutable-generation cache for incremental workflows.

    Generations are immutable JSON directories.  Promotion changes one small
    ``pointers.json`` document with ``os.replace``; this makes the logical
    current/previous switch atomic without deleting earlier generations.
    """

    def __init__(
        self,
        repo_root: Path,
        project_key: str,
        *,
        cache_root: Path | None = None,
        clinical_privacy_mode: bool = True,
    ) -> None:
        if not clinical_privacy_mode:
            # The same safe policy is used in non-clinical mode.  Keeping the
            # flag explicit prevents callers from silently losing the privacy
            # contract when a clinical project is resumed.
            clinical_privacy_mode = False
        if not _PROJECT_KEY_RE.fullmatch(project_key):
            raise CachePolicyError("project_key must be an opaque PRJ-<32 hex> value")
        self.repo_root = repo_root.resolve()
        self.cache_root = validate_cache_root(self.repo_root, cache_root)
        self.project_key = project_key
        self.project_dir = self.cache_root / project_key
        self.generations_dir = self.project_dir / "generations"
        self.pointers_path = self.project_dir / "pointers.json"
        self.lock_path = self.project_dir / ".promote.lock"
        self.clinical_privacy_mode = clinical_privacy_mode
        self.generations_dir.mkdir(parents=True, exist_ok=True)
        policy_path = self.project_dir / "cache_policy.json"
        if not policy_path.exists():
            _atomic_write_json(
                policy_path,
                {
                    "schema_version": "academic-ppt-cache-policy/1",
                    "storage_policy": "local_private_cache_only",
                    "clinical_privacy_mode": bool(clinical_privacy_mode),
                    "serialization": "json_only",
                    "project_key": project_key,
                },
            )

    @classmethod
    def from_identity(
        cls,
        repo_root: Path,
        project_identity: str,
        *,
        cache_root: Path | None = None,
        clinical_privacy_mode: bool = True,
    ) -> "ProjectCache":
        safe_root = validate_cache_root(repo_root, cache_root)
        secret = _load_or_create_key_secret(safe_root)
        project_key = derive_opaque_project_key(project_identity, secret=secret)
        return cls(
            repo_root,
            project_key,
            cache_root=safe_root,
            clinical_privacy_mode=clinical_privacy_mode,
        )

    def _read_json(self, path: Path) -> dict[str, Any]:
        try:
            value = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise CacheStateError(f"Cannot read cache JSON: {path}") from exc
        if not isinstance(value, dict):
            raise CacheStateError(f"Cache JSON must contain an object: {path}")
        return value

    def _pointers(self) -> dict[str, Any]:
        if not self.pointers_path.exists():
            return {
                "schema_version": "academic-ppt-cache-pointers/1",
                "current": None,
                "previous": None,
            }
        pointers = self._read_json(self.pointers_path)
        for name in ("current", "previous"):
            generation_id = pointers.get(name)
            if generation_id is not None and not _GENERATION_ID_RE.fullmatch(
                str(generation_id)
            ):
                raise CacheStateError(f"Invalid {name} generation pointer")
        return pointers

    def stage_generation(
        self,
        state: Mapping[str, Any],
        *,
        generation_id: str | None = None,
    ) -> str:
        validate_project_state(state)
        generation = generation_id or f"GEN-{uuid.uuid4().hex}"
        if not _GENERATION_ID_RE.fullmatch(generation):
            raise CacheStateError("generation_id must be an opaque GEN-<hex> value")
        final_dir = self.generations_dir / generation
        if final_dir.exists():
            raise CacheStateError(f"Cache generation already exists: {generation}")
        temporary_dir = Path(
            tempfile.mkdtemp(prefix=".staging-", dir=str(self.generations_dir))
        )
        try:
            _atomic_write_json(
                temporary_dir / "state.json",
                {
                    "schema_version": "academic-ppt-project-state/1",
                    "generation_id": generation,
                    "created_at": _utc_timestamp(),
                    "state": dict(state),
                },
            )
            os.replace(temporary_dir, final_dir)
        finally:
            if temporary_dir.exists():
                shutil.rmtree(temporary_dir)
        return generation

    def promote_generation(self, generation_id: str) -> dict[str, Any]:
        if not _GENERATION_ID_RE.fullmatch(str(generation_id)):
            raise CacheStateError("Invalid generation_id")
        envelope = self._read_json(
            self.generations_dir / generation_id / "state.json"
        )
        if envelope.get("generation_id") != generation_id:
            raise CacheStateError("Generation document does not match its directory")

        self.project_dir.mkdir(parents=True, exist_ok=True)
        try:
            lock_fd = os.open(
                self.lock_path,
                os.O_CREAT | os.O_EXCL | os.O_WRONLY,
                0o600,
            )
        except FileExistsError as exc:
            raise CacheStateError("Another cache promotion is in progress") from exc
        try:
            os.write(lock_fd, generation_id.encode("ascii"))
            os.fsync(lock_fd)
            pointers = self._pointers()
            current = pointers.get("current")
            if current == generation_id:
                return pointers
            promoted = {
                "schema_version": "academic-ppt-cache-pointers/1",
                "current": generation_id,
                "previous": current,
                "promoted_at": _utc_timestamp(),
            }
            _atomic_write_json(self.pointers_path, promoted)
            return promoted
        finally:
            os.close(lock_fd)
            try:
                self.lock_path.unlink()
            except FileNotFoundError:
                pass

    def commit_generation(
        self,
        state: Mapping[str, Any],
        *,
        generation_id: str | None = None,
    ) -> dict[str, Any]:
        generation = self.stage_generation(state, generation_id=generation_id)
        return self.promote_generation(generation)

    def load_generation(self, generation_id: str) -> dict[str, Any]:
        if not _GENERATION_ID_RE.fullmatch(str(generation_id)):
            raise CacheStateError("Invalid generation_id")
        envelope = self._read_json(
            self.generations_dir / generation_id / "state.json"
        )
        if envelope.get("generation_id") != generation_id:
            raise CacheStateError("Generation document does not match its directory")
        state = envelope.get("state")
        if not isinstance(state, dict):
            raise CacheStateError("Generation state must be a JSON object")
        return state

    def load_current(self) -> dict[str, Any] | None:
        generation = self._pointers().get("current")
        return self.load_generation(generation) if generation else None

    def current_generation_id(self) -> str | None:
        """Return the validated opaque current pointer without loading state."""

        generation = self._pointers().get("current")
        return str(generation) if generation else None

    def load_previous(self) -> dict[str, Any] | None:
        generation = self._pointers().get("previous")
        return self.load_generation(generation) if generation else None

    def compute_delta(self, current_manifest: Mapping[str, Any]) -> dict[str, Any]:
        """Compare a new source manifest with the current cached generation."""

        current_state = self.load_current()
        previous_manifest = (
            current_state.get("source_manifest") if current_state is not None else None
        )
        if previous_manifest is not None and not isinstance(previous_manifest, Mapping):
            raise CacheStateError("Cached source_manifest must be a JSON object")
        return compute_delta_manifest(previous_manifest, current_manifest)
