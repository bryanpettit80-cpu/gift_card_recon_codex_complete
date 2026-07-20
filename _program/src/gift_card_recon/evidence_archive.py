from __future__ import annotations

import json
import os
import shutil
import uuid
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from gift_card_recon.utils import sha256_file


class ArchiveError(RuntimeError):
    """Raised when evidence cannot be archived without risking data loss."""


@dataclass(frozen=True, slots=True)
class EvidenceItem:
    """A live close source and its intended archive classification."""

    role: str
    source_path: Path
    archive_category: str
    remove_after_publish: bool = True

    def __post_init__(self) -> None:
        role = self.role.strip()
        category = self.archive_category.strip()
        if not role:
            raise ValueError("Evidence role cannot be blank.")
        _validate_archive_category(category)
        object.__setattr__(self, "role", role)
        object.__setattr__(self, "source_path", Path(self.source_path))
        object.__setattr__(self, "archive_category", category)


@dataclass(frozen=True, slots=True)
class ArchiveRecord:
    """The immutable copy plan and audit record for one evidence source."""

    role: str
    source_path: Path
    archive_path: Path
    archive_category: str
    sha256: str
    size_bytes: int
    remove_after_publish: bool

    def __post_init__(self) -> None:
        digest = self.sha256.lower()
        if len(digest) != 64 or any(char not in "0123456789abcdef" for char in digest):
            raise ValueError("ArchiveRecord.sha256 must be a SHA-256 hex digest.")
        if self.size_bytes < 0:
            raise ValueError("ArchiveRecord.size_bytes cannot be negative.")
        object.__setattr__(self, "source_path", Path(self.source_path))
        object.__setattr__(self, "archive_path", Path(self.archive_path))
        object.__setattr__(self, "sha256", digest)


@dataclass(frozen=True, slots=True)
class ManifestArtifact:
    """Trusted identity of one published artifact recorded in a close manifest."""

    path: Path
    sha256: str
    size_bytes: int

    def __post_init__(self) -> None:
        digest = self.sha256.lower()
        if len(digest) != 64 or any(char not in "0123456789abcdef" for char in digest):
            raise ValueError("ManifestArtifact.sha256 must be a SHA-256 hex digest.")
        if self.size_bytes < 0:
            raise ValueError("ManifestArtifact.size_bytes cannot be negative.")
        object.__setattr__(self, "path", Path(self.path))
        object.__setattr__(self, "sha256", digest)


@dataclass(frozen=True, slots=True)
class CleanupResult:
    deleted_sources: tuple[Path, ...]
    pruned_directories: tuple[Path, ...]
    retained_staged_sources: tuple[Path, ...] = ()


def plan_evidence_archive(
    items: Sequence[EvidenceItem],
    *,
    archive_root: Path,
) -> list[ArchiveRecord]:
    """Create a deterministic, content-aware copy plan without changing files."""

    root = Path(archive_root)
    claimed: dict[Path, str] = {}
    records: list[ArchiveRecord] = []

    for item in items:
        source = Path(item.source_path)
        if not source.is_file():
            raise ArchiveError(f"Evidence source is missing or is not a file: {source}")
        try:
            size_bytes = source.stat().st_size
            digest = sha256_file(source)
        except OSError as exc:
            raise ArchiveError(f"Could not read evidence source {source}: {exc}") from exc

        preferred = root / Path(item.archive_category) / source.name
        destination = choose_content_safe_destination(
            preferred,
            sha256=digest,
            claimed=claimed,
        )
        records.append(
            ArchiveRecord(
                role=item.role,
                source_path=source,
                archive_path=destination,
                archive_category=item.archive_category,
                sha256=digest,
                size_bytes=size_bytes,
                remove_after_publish=item.remove_after_publish,
            )
        )

    return records


def choose_content_safe_destination(
    preferred_path: Path,
    *,
    sha256: str,
    claimed: dict[Path, str] | None = None,
) -> Path:
    """Choose an existing same-content path or a deterministic collision path."""

    preferred = Path(preferred_path)
    digest = sha256.lower()
    if len(digest) != 64 or any(char not in "0123456789abcdef" for char in digest):
        raise ValueError("sha256 must be a SHA-256 hex digest.")
    claims = claimed if claimed is not None else {}

    candidates = [preferred, preferred.with_name(f"{preferred.stem}__{digest[:12]}{preferred.suffix}")]
    candidates.extend(
        preferred.with_name(f"{preferred.stem}__{digest[:12]}_{index}{preferred.suffix}")
        for index in range(2, 1000)
    )
    for candidate in candidates:
        key = candidate.resolve(strict=False)
        claimed_digest = claims.get(key)
        if claimed_digest is not None:
            if claimed_digest == digest:
                return candidate
            continue
        if candidate.exists():
            if not candidate.is_file():
                continue
            try:
                if sha256_file(candidate) == digest:
                    claims[key] = digest
                    return candidate
            except OSError:
                continue
            continue
        claims[key] = digest
        return candidate

    raise ArchiveError(f"Could not choose a collision-safe archive path for {preferred.name}.")


def execute_archive_plan(
    records: Sequence[ArchiveRecord],
    *,
    copy_file: Callable[[Path, Path], object] = shutil.copy2,
) -> list[ArchiveRecord]:
    """Copy and verify every record without removing any source file.

    Successfully verified copies may remain if a later copy fails. Because sources
    are never deleted here, the same plan can be executed again safely.
    """

    planned = list(records)
    _preflight_sources(planned)
    completed: list[ArchiveRecord] = []

    for record in planned:
        destination = Path(record.archive_path)
        if destination.exists():
            _verify_file(destination, record.sha256, record.size_bytes, label="Archived evidence")
            completed.append(record)
            continue

        destination.parent.mkdir(parents=True, exist_ok=True)
        # Keep the atomic staging file beside its destination without repeating
        # a potentially long evidence filename. Repeating the full basename can
        # push an otherwise valid collision-safe archive path beyond Windows'
        # legacy path limit before the final os.replace occurs.
        temporary = destination.with_name(f".gc-archive-{uuid.uuid4().hex}.tmp")
        try:
            copy_file(Path(record.source_path), temporary)
            _verify_file(temporary, record.sha256, record.size_bytes, label="Staged evidence")

            # A concurrently-created destination is accepted only when its content
            # is identical. Never overwrite distinct archive evidence.
            if destination.exists():
                _verify_file(destination, record.sha256, record.size_bytes, label="Archived evidence")
            else:
                os.replace(temporary, destination)
            _verify_file(destination, record.sha256, record.size_bytes, label="Archived evidence")
            completed.append(record)
        except ArchiveError:
            raise
        except Exception as exc:
            raise ArchiveError(f"Could not archive {record.source_path} to {destination}: {exc}") from exc
        finally:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass

    return completed


def copy_and_verify_evidence(
    items: Sequence[EvidenceItem],
    *,
    archive_root: Path,
    copy_file: Callable[[Path, Path], object] = shutil.copy2,
) -> list[ArchiveRecord]:
    """Plan, copy, and hash-verify evidence while preserving all live sources."""

    records = plan_evidence_archive(items, archive_root=archive_root)
    return execute_archive_plan(records, copy_file=copy_file)


def build_close_manifest(
    *,
    store: str,
    location: str,
    period: str,
    status: str,
    source_records: Sequence[ArchiveRecord],
    artifacts: Mapping[str, Path | ManifestArtifact],
    archive_root: Path,
    generated_at: datetime | None = None,
) -> dict[str, Any]:
    """Build a serializable close manifest with source and artifact hashes."""

    timestamp = generated_at or datetime.now(timezone.utc)
    if timestamp.tzinfo is None:
        timestamp = timestamp.replace(tzinfo=timezone.utc)
    root = Path(archive_root).resolve(strict=False)

    sources: list[dict[str, Any]] = []
    for record in source_records:
        _verify_file(record.archive_path, record.sha256, record.size_bytes, label="Archived evidence")
        sources.append(
            {
                "role": record.role,
                "source_path": str(record.source_path),
                "archive_path": _relative_manifest_path(record.archive_path, root),
                "archive_category": record.archive_category,
                "sha256": record.sha256,
                "size_bytes": record.size_bytes,
                "remove_after_publish": record.remove_after_publish,
            }
        )

    artifact_rows: list[dict[str, Any]] = []
    for role, raw_artifact in sorted(artifacts.items()):
        trusted = raw_artifact if isinstance(raw_artifact, ManifestArtifact) else None
        path = trusted.path if trusted is not None else Path(raw_artifact)
        if not path.is_file():
            raise ArchiveError(f"Published artifact is missing or is not a file: {path}")
        try:
            artifact_rows.append(
                {
                    "role": str(role),
                    "path": str(path),
                    "sha256": trusted.sha256 if trusted is not None else sha256_file(path),
                    "size_bytes": trusted.size_bytes if trusted is not None else path.stat().st_size,
                }
            )
        except OSError as exc:
            raise ArchiveError(f"Could not hash published artifact {path}: {exc}") from exc

    return {
        "schema_version": 1,
        "store": str(store),
        "location": location,
        "period": period,
        "status": status,
        "generated_at": timestamp.astimezone(timezone.utc).isoformat(),
        "sources": sources,
        "artifacts": artifact_rows,
    }


def write_close_manifest_atomic(
    manifest_path: Path,
    *,
    store: str,
    location: str,
    period: str,
    status: str,
    source_records: Sequence[ArchiveRecord],
    artifacts: Mapping[str, Path | ManifestArtifact],
    archive_root: Path,
    generated_at: datetime | None = None,
    replace_file: Callable[[Path, Path], object] = os.replace,
) -> Path:
    """Atomically write a complete close manifest next to archived evidence."""

    payload = build_close_manifest(
        store=store,
        location=location,
        period=period,
        status=status,
        source_records=source_records,
        artifacts=artifacts,
        archive_root=archive_root,
        generated_at=generated_at,
    )
    path = Path(manifest_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".gc-manifest-{uuid.uuid4().hex}.tmp")
    try:
        with temporary.open("x", encoding="utf-8", newline="\n") as handle:
            json.dump(payload, handle, indent=2, ensure_ascii=False)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        replace_file(temporary, path)
    except Exception as exc:
        raise ArchiveError(f"Could not write close manifest {path}: {exc}") from exc
    finally:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass
    return path


def cleanup_after_publish(
    records: Sequence[ArchiveRecord],
    *,
    prune_period_dirs: Sequence[Path] = (),
    move_file: Callable[[Path, Path], object] = os.replace,
) -> CleanupResult:
    """Remove verified live sources, then prune only explicitly supplied trees.

    All archive and source hashes are preflighted before any source is removed. A
    changed or unverified file therefore leaves every live source intact.
    """

    removable: list[Path] = []
    seen_sources: set[Path] = set()
    for record in records:
        if not record.remove_after_publish:
            continue
        archive = Path(record.archive_path)
        source = Path(record.source_path)
        _verify_file(archive, record.sha256, record.size_bytes, label="Archived evidence")
        if source.resolve(strict=False) == archive.resolve(strict=False):
            continue
        source_key = source.resolve(strict=False)
        if source_key in seen_sources or not source.exists():
            continue
        _verify_file(source, record.sha256, record.size_bytes, label="Live evidence")
        seen_sources.add(source_key)
        removable.append(source)

    staged: list[tuple[Path, Path]] = []
    for source in removable:
        temporary = source.with_name(f".{source.name}.{uuid.uuid4().hex}.gc-cleanup")
        try:
            move_file(source, temporary)
        except OSError as exc:
            rollback_errors: list[str] = []
            for original, moved in reversed(staged):
                try:
                    os.replace(moved, original)
                except OSError as rollback_exc:
                    rollback_errors.append(f"{moved} -> {original}: {rollback_exc}")
            detail = (
                " Rollback issue(s): " + "; ".join(rollback_errors)
                if rollback_errors
                else " Every previously staged source was restored."
            )
            raise ArchiveError(
                f"Could not stage verified live evidence for cleanup {source}: {exc}.{detail}"
            ) from exc
        staged.append((source, temporary))

    # Once every source is staged, removal is post-close housekeeping. A lock
    # or transient delete failure leaves a verified hidden copy rather than
    # invalidating the published close or partially restoring live inputs.
    retained: list[Path] = []
    for _source, temporary in staged:
        try:
            temporary.unlink()
        except OSError:
            retained.append(temporary)
    deleted = [source for source, _temporary in staged]

    pruned: list[Path] = []
    for period_dir in prune_period_dirs:
        pruned.extend(_prune_empty_tree(Path(period_dir)))

    return CleanupResult(tuple(deleted), tuple(pruned), tuple(retained))


def _preflight_sources(records: Sequence[ArchiveRecord]) -> None:
    for record in records:
        _verify_file(record.source_path, record.sha256, record.size_bytes, label="Evidence source")


def _verify_file(path: Path, digest: str, size_bytes: int, *, label: str) -> None:
    candidate = Path(path)
    if not candidate.is_file():
        raise ArchiveError(f"{label} is missing or is not a file: {candidate}")
    try:
        actual_size = candidate.stat().st_size
        if actual_size != size_bytes:
            raise ArchiveError(
                f"{label} size changed for {candidate}: expected {size_bytes}, found {actual_size}."
            )
        actual_digest = sha256_file(candidate)
    except OSError as exc:
        raise ArchiveError(f"Could not verify {label.lower()} {candidate}: {exc}") from exc
    if actual_digest != digest:
        raise ArchiveError(
            f"{label} hash mismatch for {candidate}: expected {digest}, found {actual_digest}."
        )


def _relative_manifest_path(path: Path, root: Path) -> str:
    resolved = Path(path).resolve(strict=False)
    try:
        relative = resolved.relative_to(root)
    except ValueError as exc:
        raise ArchiveError(f"Archived evidence is outside the manifest archive root: {path}") from exc
    return relative.as_posix()


def _validate_archive_category(category: str) -> None:
    path = Path(category)
    if not category or path.is_absolute() or path.drive or any(part in {"", ".", ".."} for part in path.parts):
        raise ValueError(f"Archive category must be a safe relative path: {category!r}")


def _prune_empty_tree(root: Path) -> list[Path]:
    if not root.exists() or not root.is_dir() or root.is_symlink():
        return []
    removed: list[Path] = []
    descendants = sorted(
        (path for path in root.rglob("*") if path.is_dir() and not path.is_symlink()),
        key=lambda path: len(path.parts),
        reverse=True,
    )
    for directory in [*descendants, root]:
        try:
            directory.rmdir()
        except OSError:
            continue
        removed.append(directory)
    return removed
