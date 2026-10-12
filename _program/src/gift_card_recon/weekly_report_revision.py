"""Create-only report revisions that retain the original weekly evidence package."""
from __future__ import annotations

import json
import shutil
import tempfile
import uuid
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Mapping

from gift_card_recon.publication import (
    PublicationTransaction,
    file_state,
    recover_pending,
    support_root_for,
    writer_lock,
)
from gift_card_recon.utils import sha256_file
from gift_card_recon.variance_explanation import (
    EXPLANATION_SHEET_NAME,
    IDENTITY_SHEET_NAME,
    read_variance_explanation_workbook,
    verify_weekly_report_explanation_edit,
    verify_weekly_report_values,
)


REVISION_MANIFEST_NAME = "weekly_report_revision.json"


def _verified_original(package_path: Path, manifest: Mapping[str, Any]) -> Path:
    record = manifest["artifacts"]["archived_workbook"]
    relative = Path(str(record["relative_path"]))
    if relative.is_absolute() or relative.parent != Path("report"):
        raise ValueError("Original weekly workbook path is outside its report folder.")
    original = (package_path / relative).resolve()
    original.relative_to(package_path.resolve())
    if (
        not original.is_file()
        or original.stat().st_size != int(record["size_bytes"])
        or sha256_file(original) != record["sha256"]
    ):
        raise ValueError(f"Original archived weekly workbook integrity check failed: {original}")
    return original


def resolve_weekly_report_revision(
    package_path: Path, manifest: Mapping[str, Any] | None = None,
) -> Path | None:
    """Return a validated revised baseline, or None for an unrevised weekly package."""
    package_path = Path(package_path).resolve()
    revision_path = package_path / REVISION_MANIFEST_NAME
    if not revision_path.exists():
        return None
    original_manifest = package_path / "weekly_manifest.json"
    manifest = manifest or json.loads(original_manifest.read_text(encoding="utf-8"))
    revision = json.loads(revision_path.read_text(encoding="utf-8"))
    record = _validated_revision_record(revision, revision_path)
    if (
        revision.get("schema_version") != 1
        or revision.get("store") != manifest.get("store")
        or revision.get("period") != manifest.get("period")
        or revision.get("week") != manifest.get("week")
        or revision.get("original_manifest_sha256") != sha256_file(original_manifest)
    ):
        raise ValueError(f"Weekly report revision identity or original manifest binding is invalid: {revision_path}")
    original = _verified_original(package_path, manifest)
    if revision.get("original_report_sha256") != sha256_file(original):
        raise ValueError("Weekly report revision does not match the original archived report.")
    relative = Path(str(record["relative_path"]))
    if (
        relative.is_absolute() or len(relative.parts) != 3
        or relative.parts[0] != "report-revisions" or relative.parts[1] in {".", ".."}
        or relative.name != original.name
    ):
        raise ValueError("Revised weekly workbook path is outside its revision folder.")
    revised = (package_path / relative).resolve()
    revised.relative_to(package_path)
    if (
        not revised.is_file() or revised.stat().st_size != record["size_bytes"]
        or sha256_file(revised) != record["sha256"]
    ):
        raise ValueError(f"Revised weekly workbook integrity check failed: {revised}")
    read_variance_explanation_workbook(
        revised, expected_store=str(manifest["store"]),
        expected_week_start=date.fromisoformat(manifest["week"]["start"]),
        expected_week_end=date.fromisoformat(manifest["week"]["end"]), require_text=False,
    )
    return revised


def _validated_revision_record(
    revision: Any, revision_path: Path,
) -> Mapping[str, Any]:
    """Validate revision objects before their values reach path or size operations."""
    if not isinstance(revision, Mapping):
        raise ValueError(f"Weekly report revision manifest structure is invalid: {revision_path}")
    record = revision.get("revised_workbook")
    if not isinstance(record, Mapping):
        raise ValueError(f"Weekly report revision manifest structure is invalid: {revision_path}")
    relative_path = record.get("relative_path")
    size_bytes = record.get("size_bytes")
    digest = record.get("sha256")
    if (
        not isinstance(relative_path, str)
        or not relative_path
        or not isinstance(size_bytes, int)
        or isinstance(size_bytes, bool)
        or size_bytes < 0
        or not isinstance(digest, str)
        or not digest
    ):
        raise ValueError(f"Weekly report revision manifest structure is invalid: {revision_path}")
    return record


def publish_weekly_report_revision(
    package_path: Path, canonical_path: Path, revised_workbook_path: Path,
    *, expected_canonical_sha256: str | None = None,
) -> Path:
    """Publish one report upgrade without replacing original evidence or user edits.

    This is deliberately create-only. Subsequent revisions require a reviewed new
    revision chain; silently replacing the first revision is not supported.
    """
    support_root = support_root_for(Path(package_path))
    with writer_lock(support_root):
        recover_pending(support_root)
        return _publish_weekly_report_revision(
            package_path, canonical_path, revised_workbook_path,
            expected_canonical_sha256=expected_canonical_sha256,
            support_root=support_root,
        )


def _publish_weekly_report_revision(
    package_path: Path, canonical_path: Path, revised_workbook_path: Path,
    *, expected_canonical_sha256: str | None, support_root: Path,
) -> Path:
    package_path = Path(package_path).resolve()
    canonical_path = Path(canonical_path).resolve()
    revised_workbook_path = Path(revised_workbook_path).resolve()
    sidecar = package_path / REVISION_MANIFEST_NAME
    if sidecar.exists():
        raise FileExistsError(f"Weekly report revision already exists: {sidecar}")
    manifest_path = package_path / "weekly_manifest.json"
    manifest_state = file_state(manifest_path)
    manifest_digest = sha256_file(manifest_path)
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    if manifest.get("schema_version") != 1:
        raise ValueError("Original weekly manifest schema version is unsupported.")
    original = _verified_original(package_path, manifest)
    original_digest = sha256_file(original)
    if canonical_path == original or canonical_path.is_relative_to(package_path):
        raise ValueError("Editable weekly report must be outside its archived evidence package.")
    if canonical_path.name != original.name or not canonical_path.is_file():
        raise ValueError("Current weekly report differs from its original archive; preserve and review its edits before revision.")
    canonical_digest = sha256_file(canonical_path)
    canonical_state = "file:" + canonical_digest
    if expected_canonical_sha256 is not None and canonical_digest != expected_canonical_sha256:
        raise ValueError("Current weekly report changed since the reviewed snapshot.")
    from openpyxl import load_workbook

    original_workbook = load_workbook(original, read_only=True, keep_links=False)
    try:
        has_embedded_explanation = (
            EXPLANATION_SHEET_NAME in original_workbook.sheetnames
            or IDENTITY_SHEET_NAME in original_workbook.sheetnames
        )
    finally:
        original_workbook.close()
    if canonical_digest != original_digest:
        if expected_canonical_sha256 is None:
            raise ValueError("Current weekly report differs from its original archive; preserve and review its edits before revision.")
        if has_embedded_explanation:
            verify_weekly_report_explanation_edit(canonical_path, original)
        else:
            verify_weekly_report_values(canonical_path, original)
    candidate_explanation = read_variance_explanation_workbook(
        revised_workbook_path, expected_store=str(manifest["store"]),
        expected_week_start=date.fromisoformat(manifest["week"]["start"]),
        expected_week_end=date.fromisoformat(manifest["week"]["end"]), require_text=False,
    )
    if has_embedded_explanation:
        canonical_explanation = read_variance_explanation_workbook(
            canonical_path, expected_store=str(manifest["store"]),
            expected_week_start=date.fromisoformat(manifest["week"]["start"]),
            expected_week_end=date.fromisoformat(manifest["week"]["end"]), require_text=False,
        )
        if candidate_explanation.explanation != canonical_explanation.explanation:
            raise ValueError("Revised weekly report must retain the current operator explanation.")
    verify_weekly_report_explanation_edit(revised_workbook_path, revised_workbook_path)
    source_digest = sha256_file(revised_workbook_path)
    revision_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ") + "-" + uuid.uuid4().hex[:8]
    revision_dir = package_path / "report-revisions" / revision_id
    revised = revision_dir / original.name
    with tempfile.TemporaryDirectory(prefix="weekly-revision-", dir=support_root) as temporary:
        staged_workbook = Path(temporary) / original.name
        staged_sidecar = Path(temporary) / REVISION_MANIFEST_NAME
        shutil.copy2(revised_workbook_path, staged_workbook)
        if sha256_file(staged_workbook) != source_digest or sha256_file(revised_workbook_path) != source_digest:
            raise ValueError("Revised weekly workbook changed during staging.")
        revision = {
            "schema_version": 1, "store": manifest["store"], "period": manifest["period"],
            "week": manifest["week"], "created_at": datetime.now(timezone.utc).isoformat(),
            "original_manifest_sha256": manifest_digest, "original_report_sha256": original_digest,
            "revised_workbook": {
                "relative_path": revised.relative_to(package_path).as_posix(),
                "size_bytes": staged_workbook.stat().st_size, "sha256": source_digest,
            },
        }
        staged_sidecar.write_text(json.dumps(revision, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        if (
            sha256_file(manifest_path) != manifest_digest
            or sha256_file(original) != original_digest
            or sha256_file(canonical_path) != canonical_digest
        ):
            raise ValueError("Weekly evidence or current report changed during revision staging.")
        if sidecar.exists():
            raise FileExistsError(f"Weekly report revision already exists: {sidecar}")
        if revision_dir.exists():
            raise FileExistsError(f"Weekly report revision directory already exists: {revision_dir}")
        with PublicationTransaction(support_root) as transaction:
            transaction.replace(staged_workbook, revised, expected_before=None)
            if sha256_file(canonical_path) != canonical_digest:
                raise ValueError("Current weekly report changed before revision publication.")
            transaction.replace(staged_workbook, canonical_path, expected_before=canonical_state)
            if sidecar.exists():
                raise FileExistsError(f"Weekly report revision already exists: {sidecar}")
            transaction.replace(staged_sidecar, sidecar, expected_before=None)
            resolved = resolve_weekly_report_revision(package_path, manifest)
            if (resolved != revised or sha256_file(canonical_path) != source_digest
                    or file_state(manifest_path) != manifest_state):
                raise ValueError("Weekly report revision verification failed after publication.")
            transaction.commit()
        return sidecar
