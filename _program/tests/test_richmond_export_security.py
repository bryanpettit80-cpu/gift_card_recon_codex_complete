from pathlib import Path


def test_richmond_scheduled_task_runs_protected_local_copy():
    installer = (
        Path(__file__).resolve().parents[1]
        / "support"
        / "richmond_micros_export"
        / "Install-DailyGiftCardCopyTask.cmd"
    )
    text = installer.read_text(encoding="utf-8")

    assert 'set "INSTALL_DIR=%ProgramData%\\GiftCardRecon\\RichmondMicrosExport"' in text
    assert 'copy /Y "%SOURCE_SCRIPT%" "%SCRIPT_PATH%"' in text
    assert 'icacls "%INSTALL_DIR%" /inheritance:r' in text

    create_line = next(line for line in text.splitlines() if line.startswith("schtasks /Create"))
    assert "%SCRIPT_PATH%" in create_line
    assert "Dropbox" not in create_line
