from pathlib import Path


def test_program_repo_does_not_ship_legacy_business_placeholders() -> None:
    repository_root = Path(__file__).resolve().parents[2]
    legacy_placeholders = (
        repository_root / "9354 - Weekly" / "activity" / ".gitkeep",
        repository_root / "9355 - Weekly" / "activity" / ".gitkeep",
        repository_root / "Monthly Close" / ".gitkeep",
        repository_root / "Monthly Close" / "9354" / ".gitkeep",
        repository_root / "Monthly Close" / "9355" / ".gitkeep",
        repository_root / "Output" / ".gitkeep",
        repository_root / "Archive - Old Files" / ".gitkeep",
    )

    assert not [path for path in legacy_placeholders if path.exists()]
