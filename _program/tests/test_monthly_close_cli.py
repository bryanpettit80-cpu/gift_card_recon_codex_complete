from __future__ import annotations

from datetime import date
from decimal import Decimal
import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from gift_card_recon.close_assessment import ControlDisposition, ControlOutcome, build_close_assessment
from gift_card_recon.fiscal_calendar import fiscal_period_for_label
from gift_card_recon.models import DardenCreditMemo
from gift_card_recon.monthly_close_cli import (
    CloseJob,
    ArchiveReissue,
    SHARED_DARDEN_INBOX,
    build_parser,
    discover_close_jobs,
    main,
    _resolve_input_dir,
    _resolve_archive_reissue,
)
from gift_card_recon.monthly_close_service import CloseBlockedError
from gift_card_recon.parsers import ParseError
from gift_card_recon.utils import sha256_file


def test_no_argument_parser_uses_shared_inbox_mode() -> None:
    args = build_parser().parse_args([])

    assert args.store is None
    assert args.period is None
    assert args.darden_path is None
    assert SHARED_DARDEN_INBOX == "Darden Reports - Drop Here"


def test_archive_reissue_requires_identity_and_rejects_manual_evidence_paths(
    tmp_path: Path,
    capsys,
) -> None:
    assert main(["--reissue-from-archive", "--archive-root", str(tmp_path)]) == 1
    assert "requires both --store and --period" in capsys.readouterr().out

    assert main(
        [
            "--reissue-from-archive",
            "--store",
            "9355",
            "--period",
            "FY27-M01",
            "--input-dir",
            str(tmp_path / "manual"),
        ]
    ) == 1
    assert "cannot be combined with --input-dir" in capsys.readouterr().out


def test_archive_reissue_verifies_manifest_sources_and_forces_safe_run_flags(
    tmp_path: Path,
    monkeypatch,
) -> None:
    period = fiscal_period_for_label("FY27-M01")
    archive_root = tmp_path / "Archive - Old Files"
    package = archive_root / "Monthly Close" / "9355" / period.folder_name
    paths = [
        ("Gift Card Summary", package / "summary" / "07.05.2026 9355 Gift Card Summary.xlsx"),
        *[
            ("Weekly Gift Card Activity", package / "activity" / f"{week:%m.%d.%Y} 9355 Gift Card Activity.xls")
            for week in period.expected_week_endings
        ],
        ("Darden Credit Memo", package / "darden" / "memo.pdf"),
        ("Micros Daily System Totals", package / "micros" / "DLYSYSTT.TXT"),
        ("Micros Tender Detail", package / "micros" / "TENDER_DETAIL.TXT"),
    ]
    rows = []
    for index, (role, path) in enumerate(paths):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(f"source-{index}".encode())
        rows.append(
            {
                "role": role,
                "archive_path": path.relative_to(archive_root).as_posix(),
                "sha256": sha256_file(path),
                "size_bytes": path.stat().st_size,
            }
        )
    (package / "close_manifest.json").write_text(
        json.dumps({"store": "9355", "period": period.period_key, "sources": rows}),
        encoding="utf-8",
    )
    darden_path = package / "darden" / "memo.pdf"
    monkeypatch.setattr(
        "gift_card_recon.monthly_close_cli.parse_darden_credit_memo",
        lambda path: _report(Path(path), "9355"),
    )

    resolved = _resolve_archive_reissue(
        archive_root=archive_root,
        store="9355",
        period=period.period_key,
    )
    assert resolved.input_dir == package.resolve(strict=False)
    assert resolved.micros_path == package.resolve(strict=False) / "micros"
    assert resolved.job.darden_path == darden_path.resolve(strict=False)

    monkeypatch.setattr(
        "gift_card_recon.monthly_close_cli._resolve_archive_reissue",
        lambda **_kwargs: resolved,
    )
    calls: list[dict[str, object]] = []
    monkeypatch.setattr(
        "gift_card_recon.monthly_close_cli.run_monthly_close_service",
        lambda **kwargs: calls.append(kwargs) or SimpleNamespace(),
    )
    monkeypatch.setattr("gift_card_recon.monthly_close_cli._print_success", lambda _result: None)

    assert main(
        [
            "--reissue-from-archive",
            "--store",
            "9355",
            "--period",
            period.period_key,
            "--archive-root",
            str(archive_root),
            "--output-file",
            str(tmp_path / "reissue.xlsx"),
        ]
    ) == 0
    assert calls[0]["input_dir"] == resolved.input_dir
    assert calls[0]["micros_path"] == resolved.micros_path
    assert calls[0]["cleanup_sources"] is False
    assert calls[0]["allow_unconfigured_micros"] is True

    paths[0][1].write_bytes(b"tampered")
    with pytest.raises(ParseError, match="hash does not match|size does not match"):
        _resolve_archive_reissue(
            archive_root=archive_root,
            store="9355",
            period=period.period_key,
        )


def test_shared_inbox_derives_store_and_period_from_every_pdf(tmp_path: Path, monkeypatch) -> None:
    inbox = tmp_path / SHARED_DARDEN_INBOX
    inbox.mkdir()
    richmond_path = inbox / "richmond.pdf"
    beach_path = inbox / "beach.pdf"
    richmond_path.write_bytes(b"pdf")
    beach_path.write_bytes(b"pdf")
    reports = {
        richmond_path: _report(richmond_path, "9354"),
        beach_path: _report(beach_path, "9355"),
    }
    monkeypatch.setattr(
        "gift_card_recon.monthly_close_cli.parse_darden_credit_memo",
        lambda path: reports[Path(path)],
    )

    jobs, errors = discover_close_jobs(
        inbox=inbox,
        explicit_darden=None,
        store=None,
        period=None,
    )

    assert errors == []
    assert [(job.store, job.fiscal_period.period_key) for job in jobs] == [
        ("9354", "FY27-M01"),
        ("9355", "FY27-M01"),
    ]


def test_duplicate_store_period_pdfs_are_rejected_without_hiding_either(tmp_path: Path, monkeypatch) -> None:
    inbox = tmp_path / SHARED_DARDEN_INBOX
    inbox.mkdir()
    paths = [inbox / "first.pdf", inbox / "second.pdf"]
    for path in paths:
        path.write_bytes(b"pdf")
    monkeypatch.setattr(
        "gift_card_recon.monthly_close_cli.parse_darden_credit_memo",
        lambda path: _report(Path(path), "9355"),
    )

    jobs, errors = discover_close_jobs(
        inbox=inbox,
        explicit_darden=None,
        store=None,
        period=None,
    )

    assert jobs == []
    assert len(errors) == 1
    assert "first.pdf" in str(errors[0]) and "second.pdf" in str(errors[0])
    assert errors[0].store == "9355"
    assert errors[0].period == "FY27-M01"


def test_shared_inbox_runs_locations_independently(tmp_path: Path, monkeypatch) -> None:
    period = fiscal_period_for_label("FY27-M01")
    jobs = [
        CloseJob("9354", period, tmp_path / "richmond.pdf", _report(tmp_path / "richmond.pdf", "9354")),
        CloseJob("9355", period, tmp_path / "beach.pdf", _report(tmp_path / "beach.pdf", "9355")),
    ]
    monkeypatch.setattr(
        "gift_card_recon.monthly_close_cli.discover_close_jobs",
        lambda **_kwargs: (jobs, []),
    )
    monkeypatch.setattr(
        "gift_card_recon.monthly_close_cli._resolve_input_dir",
        lambda job, **_kwargs: tmp_path / job.store,
    )
    calls: list[str] = []

    def run(**kwargs):
        calls.append(kwargs["store"])
        if kwargs["store"] == "9354":
            assessment = build_close_assessment(
                store="9354",
                darden_variance=Decimal("0.00"),
                controls=(
                    ControlOutcome(
                        "simulated_failure",
                        "Simulated failure",
                        ControlDisposition.BLOCK,
                        "Richmond failed independently.",
                    ),
                ),
            )
            raise CloseBlockedError("Richmond failed independently.", assessment=assessment)
        return SimpleNamespace()

    monkeypatch.setattr("gift_card_recon.monthly_close_cli.run_monthly_close_service", run)
    monkeypatch.setattr("gift_card_recon.monthly_close_cli._print_success", lambda _result: None)

    exit_code = main(
        [
            "--input-root",
            str(tmp_path / "Monthly Close"),
            "--archive-root",
            str(tmp_path / "Archive - Old Files"),
            "--output-dir",
            str(tmp_path / "Output"),
        ]
    )

    assert exit_code == 1
    assert calls == ["9354", "9355"]


def test_explicit_store_filters_other_valid_inbox_reports_without_error(tmp_path: Path, monkeypatch) -> None:
    inbox = tmp_path / SHARED_DARDEN_INBOX
    inbox.mkdir()
    paths = [inbox / "richmond.pdf", inbox / "beach.pdf"]
    for path in paths:
        path.write_bytes(b"pdf")
    reports = {
        paths[0]: _report(paths[0], "9354"),
        paths[1]: _report(paths[1], "9355"),
    }
    monkeypatch.setattr(
        "gift_card_recon.monthly_close_cli.parse_darden_credit_memo",
        lambda path: reports[Path(path)],
    )

    jobs, errors = discover_close_jobs(
        inbox=inbox,
        explicit_darden=None,
        store="9355",
        period="FY27-M01",
    )

    assert errors == []
    assert [job.store for job in jobs] == ["9355"]


def test_archived_inputs_are_used_only_with_explicit_input_dir(tmp_path: Path) -> None:
    period = fiscal_period_for_label("FY27-M01")
    job = CloseJob("9355", period, tmp_path / "memo.pdf", _report(tmp_path / "memo.pdf", "9355"))
    input_root = tmp_path / "Monthly Close"
    archive_root = tmp_path / "Archive - Old Files"
    archived = archive_root / "Monthly Close" / "9355" / period.folder_name
    (archived / "summary").mkdir(parents=True)
    (archived / "activity").mkdir()
    (archived / "summary" / "07.05.2026 9355 Gift Card Summary.xlsx").write_bytes(b"summary")
    (archived / "activity" / "07.05.2026 9355 Gift Card Activity.xlsx").write_bytes(b"activity")

    resolved = _resolve_input_dir(
        job,
        input_root=input_root,
        archive_root=archive_root,
        explicit_input=None,
        stage_weekly=False,
    )

    assert resolved == input_root / "9355 Virginia Beach" / period.folder_name
    assert resolved != archived


def test_explicit_legacy_lowercase_archive_remains_readable(tmp_path: Path) -> None:
    period = fiscal_period_for_label("FY27-M01")
    job = CloseJob("9355", period, tmp_path / "memo.pdf", _report(tmp_path / "memo.pdf", "9355"))
    legacy = tmp_path / "Archive - Old Files" / "monthly-close" / "9355" / period.folder_name
    legacy.mkdir(parents=True)

    resolved = _resolve_input_dir(
        job,
        input_root=tmp_path / "Monthly Close",
        archive_root=tmp_path / "Archive - Old Files",
        explicit_input=legacy,
        stage_weekly=False,
    )

    assert resolved == legacy


def test_prepare_only_uses_strict_assessment_and_returns_nonzero_when_blocked(
    tmp_path: Path,
    monkeypatch,
) -> None:
    period = fiscal_period_for_label("FY27-M01")
    job = CloseJob("9355", period, tmp_path / "memo.pdf", _report(tmp_path / "memo.pdf", "9355"))
    blocked = build_close_assessment(
        store="9355",
        darden_variance=Decimal("0.00"),
        controls=(
            ControlOutcome(
                "tender_evidence",
                "Tender evidence",
                ControlDisposition.BLOCK,
                "Tender evidence is missing.",
            ),
        ),
    )
    monkeypatch.setattr(
        "gift_card_recon.monthly_close_cli.discover_close_jobs",
        lambda **_kwargs: ([job], []),
    )
    monkeypatch.setattr(
        "gift_card_recon.monthly_close_cli._resolve_input_dir",
        lambda *_args, **_kwargs: tmp_path / "input",
    )
    monkeypatch.setattr(
        "gift_card_recon.monthly_close_cli.assess_monthly_close_inputs",
        lambda **_kwargs: blocked,
    )

    exit_code = main(
        [
            "--prepare-only",
            "--input-root",
            str(tmp_path / "Monthly Close"),
            "--archive-root",
            str(tmp_path / "Archive - Old Files"),
        ]
    )

    assert exit_code == 1


def _report(path: Path, store: str) -> DardenCreditMemo:
    return DardenCreditMemo(
        source_file=path,
        store=store,
        period_start=date(2026, 6, 1),
        period_end=date(2026, 7, 5),
        total=Decimal("-200.00"),
    )
