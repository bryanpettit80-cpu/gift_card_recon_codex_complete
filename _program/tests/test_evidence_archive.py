from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

import pytest

from gift_card_recon.evidence_archive import (
    ArchiveError,
    EvidenceItem,
    build_close_manifest,
    cleanup_after_publish,
    copy_and_verify_evidence,
    execute_archive_plan,
    plan_evidence_archive,
    write_close_manifest_atomic,
)
from gift_card_recon.utils import sha256_file


def test_archive_collision_uses_content_hash_and_preserves_both_files(tmp_path: Path) -> None:
    first = tmp_path / "live" / "summary" / "report.xlsx"
    second = tmp_path / "live" / "activity" / "report.xlsx"
    first.parent.mkdir(parents=True)
    second.parent.mkdir(parents=True)
    first.write_bytes(b"first evidence")
    second.write_bytes(b"second evidence")
    staging_names: list[str] = []

    def capture_staging_name(source: Path, destination: Path) -> None:
        staging_names.append(destination.name)
        destination.write_bytes(source.read_bytes())

    records = copy_and_verify_evidence(
        [
            EvidenceItem("summary", first, "evidence"),
            EvidenceItem("activity", second, "evidence"),
        ],
        archive_root=tmp_path / "archive",
        copy_file=capture_staging_name,
    )

    assert len(staging_names) == 2
    assert all(name.startswith(".gc-archive-") and name.endswith(".tmp") for name in staging_names)
    assert all("report" not in name for name in staging_names)
    assert records[0].archive_path.name == "report.xlsx"
    assert records[1].archive_path.name == f"report__{records[1].sha256[:12]}.xlsx"
    assert records[0].archive_path.read_bytes() == b"first evidence"
    assert records[1].archive_path.read_bytes() == b"second evidence"
    assert first.exists()
    assert second.exists()


def test_archive_plan_is_idempotent_for_existing_verified_copies(tmp_path: Path) -> None:
    source = tmp_path / "live" / "memo.pdf"
    source.parent.mkdir()
    source.write_bytes(b"Darden evidence")
    item = EvidenceItem("darden_credit_memo", source, "darden")

    first = copy_and_verify_evidence([item], archive_root=tmp_path / "archive")
    archived_mtime = first[0].archive_path.stat().st_mtime_ns
    second_plan = plan_evidence_archive([item], archive_root=tmp_path / "archive")
    second = execute_archive_plan(second_plan)

    assert second == first
    assert second[0].archive_path.stat().st_mtime_ns == archived_mtime
    assert source.exists()


def test_partial_copy_failure_preserves_every_live_source(tmp_path: Path) -> None:
    sources = [tmp_path / "live" / f"source-{index}.txt" for index in range(2)]
    for index, source in enumerate(sources):
        source.parent.mkdir(parents=True, exist_ok=True)
        source.write_text(f"evidence {index}", encoding="utf-8")
    records = plan_evidence_archive(
        [EvidenceItem(f"source_{index}", source, "raw") for index, source in enumerate(sources)],
        archive_root=tmp_path / "archive",
    )
    calls = 0

    def fail_second_copy(source: Path, destination: Path) -> None:
        nonlocal calls
        calls += 1
        if calls == 2:
            raise OSError("synthetic copy failure")
        destination.write_bytes(source.read_bytes())

    with pytest.raises(ArchiveError, match="synthetic copy failure"):
        execute_archive_plan(records, copy_file=fail_second_copy)

    assert all(source.exists() for source in sources)
    assert records[0].archive_path.exists()
    assert not records[1].archive_path.exists()
    assert not list((tmp_path / "archive").rglob("*.tmp"))


def test_cleanup_preflights_all_hashes_before_deleting_any_source(tmp_path: Path) -> None:
    first = tmp_path / "period" / "summary" / "first.xlsx"
    second = tmp_path / "period" / "darden" / "second.pdf"
    first.parent.mkdir(parents=True)
    second.parent.mkdir(parents=True)
    first.write_bytes(b"summary")
    second.write_bytes(b"memo")
    records = copy_and_verify_evidence(
        [EvidenceItem("summary", first, "summary"), EvidenceItem("darden", second, "darden")],
        archive_root=tmp_path / "archive",
    )
    records[1].archive_path.write_bytes(b"corrupt")

    with pytest.raises(ArchiveError, match="Archived evidence"):
        cleanup_after_publish(records, prune_period_dirs=[tmp_path / "period"])

    assert first.exists()
    assert second.exists()
    assert (tmp_path / "period").exists()


def test_cleanup_removes_only_verified_sources_and_prunes_explicit_period_tree(tmp_path: Path) -> None:
    period = tmp_path / "Monthly Close" / "9354" / "FY27 M01 - Fiscal June"
    removable = period / "summary" / "summary.xlsx"
    retained = period / "micros" / "DLYSYSTT.TXT"
    removable.parent.mkdir(parents=True)
    retained.parent.mkdir(parents=True)
    removable.write_bytes(b"summary")
    retained.write_bytes(b"system totals")
    records = copy_and_verify_evidence(
        [
            EvidenceItem("summary", removable, "summary"),
            EvidenceItem("system_totals", retained, "micros", remove_after_publish=False),
        ],
        archive_root=tmp_path / "archive",
    )

    result = cleanup_after_publish(records, prune_period_dirs=[period])

    assert result.deleted_sources == (removable,)
    assert not removable.exists()
    assert retained.exists()
    assert period.exists()
    assert not (period / "summary").exists()
    assert (period / "micros").exists()


def test_cleanup_staging_failure_restores_every_live_source(tmp_path: Path) -> None:
    sources = [tmp_path / "period" / "summary.xlsx", tmp_path / "period" / "memo.pdf"]
    sources[0].parent.mkdir(parents=True)
    for index, source in enumerate(sources):
        source.write_bytes(f"source-{index}".encode())
    records = copy_and_verify_evidence(
        [
            EvidenceItem("summary", sources[0], "summary"),
            EvidenceItem("darden", sources[1], "darden"),
        ],
        archive_root=tmp_path / "archive",
    )
    calls = 0

    def fail_second_move(source: Path, destination: Path):
        nonlocal calls
        calls += 1
        if calls == 2:
            raise OSError("simulated live-file lock")
        source.replace(destination)

    with pytest.raises(ArchiveError, match="Every previously staged source was restored"):
        cleanup_after_publish(records, move_file=fail_second_move)

    assert all(source.exists() for source in sources)
    assert not list((tmp_path / "period").glob("*.gc-cleanup"))


def test_manifest_records_relative_archive_paths_and_artifact_hashes(tmp_path: Path) -> None:
    source = tmp_path / "live" / "summary.xlsx"
    artifact = tmp_path / "Output" / "Richmond_9354_FY27-M01_Monthly_Close.xlsx"
    source.parent.mkdir()
    artifact.parent.mkdir()
    source.write_bytes(b"summary")
    artifact.write_bytes(b"workbook")
    archive_root = tmp_path / "Archive - Old Files" / "Monthly Close" / "9354" / "FY27 M01 - Fiscal June"
    records = copy_and_verify_evidence(
        [EvidenceItem("summary", source, "summary")],
        archive_root=archive_root,
    )
    generated_at = datetime(2026, 7, 6, 12, 30, tzinfo=timezone.utc)

    manifest = build_close_manifest(
        store="9354",
        location="Richmond",
        period="FY27-M01",
        status="CLOSED WITH REVIEW",
        source_records=records,
        artifacts={"workbook": artifact},
        archive_root=archive_root,
        generated_at=generated_at,
    )

    assert manifest["store"] == "9354"
    assert manifest["location"] == "Richmond"
    assert manifest["sources"][0]["archive_path"] == "summary/summary.xlsx"
    assert manifest["sources"][0]["sha256"] == sha256_file(source)
    assert manifest["artifacts"][0]["sha256"] == sha256_file(artifact)
    assert manifest["generated_at"] == "2026-07-06T12:30:00+00:00"


def test_manifest_write_is_atomic_and_valid_json(tmp_path: Path) -> None:
    source = tmp_path / "live" / "memo.pdf"
    artifact = tmp_path / "output" / "report.pdf"
    source.parent.mkdir()
    artifact.parent.mkdir()
    source.write_bytes(b"memo")
    artifact.write_bytes(b"pdf")
    archive_root = tmp_path / "archive"
    records = copy_and_verify_evidence(
        [EvidenceItem("darden", source, "darden")],
        archive_root=archive_root,
    )
    manifest_path = archive_root / "close-manifest.json"
    manifest_staging_names: list[str] = []

    def capture_manifest_replace(source: Path, destination: Path) -> None:
        manifest_staging_names.append(source.name)
        source.replace(destination)

    written = write_close_manifest_atomic(
        manifest_path,
        store="9355",
        location="Virginia Beach",
        period="FY27-M01",
        status="CLOSED",
        source_records=records,
        artifacts={"pdf": artifact},
        archive_root=archive_root,
        generated_at=datetime(2026, 7, 6, tzinfo=timezone.utc),
        replace_file=capture_manifest_replace,
    )

    assert written == manifest_path
    assert len(manifest_staging_names) == 1
    assert manifest_staging_names[0].startswith(".gc-manifest-")
    assert manifest_staging_names[0].endswith(".tmp")
    assert manifest_path.name not in manifest_staging_names[0]
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    assert payload["status"] == "CLOSED"
    assert payload["sources"][0]["archive_path"] == "darden/memo.pdf"
    assert not list(archive_root.glob(".*.tmp"))
