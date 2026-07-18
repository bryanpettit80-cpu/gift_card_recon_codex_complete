from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]
SUPPORT_ROOT = REPO_ROOT / "_program" / "support" / "richmond_micros_export"


def test_daily_copy_task_uses_verified_local_snapshot(tmp_path: Path) -> None:
    synced_setup = tmp_path / "setup"
    synced_setup.mkdir()
    installer = synced_setup / "Install-DailyGiftCardCopyTask.cmd"
    source_script = synced_setup / "Copy-GiftCardExportToDropbox.cmd"
    shutil.copy2(SUPPORT_ROOT / installer.name, installer)
    shutil.copy2(SUPPORT_ROOT / source_script.name, source_script)

    stub_bin = tmp_path / "stub-bin"
    stub_bin.mkdir()
    schedule_capture = tmp_path / "scheduled-action.txt"
    (stub_bin / "schtasks.cmd").write_text(
        '@echo off\n>> "%SCHEDULE_CAPTURE%" echo %*\nexit /b 0\n',
        encoding="utf-8",
    )

    local_app_data = tmp_path / "Local App Data"
    env = os.environ.copy()
    env["LOCALAPPDATA"] = str(local_app_data)
    env["SCHEDULE_CAPTURE"] = str(schedule_capture)
    env["PATH"] = f"{stub_bin}{os.pathsep}{env['PATH']}"

    completed = subprocess.run(
        [env.get("COMSPEC", "cmd.exe"), "/d", "/c", installer.name],
        cwd=synced_setup,
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stdout + completed.stderr
    installed_script = (
        local_app_data
        / "GiftCardRecon"
        / "RichmondMicrosExport"
        / source_script.name
    )
    assert installed_script.read_bytes() == source_script.read_bytes()

    schedule_calls = schedule_capture.read_text(encoding="utf-8").splitlines()
    assert len(schedule_calls) == 2
    safe_action, scheduled_action = schedule_calls
    assert "Gift Card Export Copy to Dropbox" in safe_action
    assert "cmd.exe" in safe_action.lower()
    assert "/d /c exit 0" in safe_action.lower()
    assert source_script.name not in safe_action
    assert "/SC DAILY /ST 06:35" in safe_action
    assert "Gift Card Export Copy to Dropbox" in scheduled_action
    assert str(installed_script) in scheduled_action
    assert str(source_script) not in scheduled_action
    assert "/SC DAILY /ST 06:35" in scheduled_action

    installed_bytes = installed_script.read_bytes()
    source_script.write_text("@echo malicious replacement\n", encoding="utf-8")
    assert installed_script.read_bytes() == installed_bytes

    installer_text = installer.read_text(encoding="utf-8")
    assert "[System.Security.Cryptography.SHA256]::Create()" in installer_text


def test_daily_copy_task_fails_closed_when_source_is_missing(tmp_path: Path) -> None:
    synced_setup = tmp_path / "setup"
    synced_setup.mkdir()
    installer = synced_setup / "Install-DailyGiftCardCopyTask.cmd"
    shutil.copy2(SUPPORT_ROOT / installer.name, installer)

    stub_bin = tmp_path / "stub-bin"
    stub_bin.mkdir()
    schedule_capture = tmp_path / "scheduled-action.txt"
    (stub_bin / "schtasks.cmd").write_text(
        '@echo off\n>> "%SCHEDULE_CAPTURE%" echo %*\nexit /b 0\n',
        encoding="utf-8",
    )

    env = os.environ.copy()
    env["LOCALAPPDATA"] = str(tmp_path / "Local App Data")
    env["SCHEDULE_CAPTURE"] = str(schedule_capture)
    env["PATH"] = f"{stub_bin}{os.pathsep}{env['PATH']}"

    completed = subprocess.run(
        [env.get("COMSPEC", "cmd.exe"), "/d", "/c", installer.name],
        cwd=synced_setup,
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 20, completed.stdout + completed.stderr
    schedule_calls = schedule_capture.read_text(encoding="utf-8").splitlines()
    assert len(schedule_calls) == 1
    safe_action = schedule_calls[0]
    assert "Gift Card Export Copy to Dropbox" in safe_action
    assert "cmd.exe" in safe_action.lower()
    assert "/d /c exit 0" in safe_action.lower()
    assert "Copy-GiftCardExportToDropbox.cmd" not in safe_action
    assert "/SC DAILY /ST 06:35" in safe_action


def test_daily_copy_task_stops_when_neutralization_fails(tmp_path: Path) -> None:
    synced_setup = tmp_path / "setup"
    synced_setup.mkdir()
    installer = synced_setup / "Install-DailyGiftCardCopyTask.cmd"
    source_script = synced_setup / "Copy-GiftCardExportToDropbox.cmd"
    shutil.copy2(SUPPORT_ROOT / installer.name, installer)
    shutil.copy2(SUPPORT_ROOT / source_script.name, source_script)

    stub_bin = tmp_path / "stub-bin"
    stub_bin.mkdir()
    schedule_capture = tmp_path / "scheduled-action.txt"
    (stub_bin / "schtasks.cmd").write_text(
        '@echo off\n>> "%SCHEDULE_CAPTURE%" echo %*\nexit /b 1\n',
        encoding="utf-8",
    )

    local_app_data = tmp_path / "Local App Data"
    env = os.environ.copy()
    env["LOCALAPPDATA"] = str(local_app_data)
    env["SCHEDULE_CAPTURE"] = str(schedule_capture)
    env["PATH"] = f"{stub_bin}{os.pathsep}{env['PATH']}"

    completed = subprocess.run(
        [env.get("COMSPEC", "cmd.exe"), "/d", "/c", installer.name],
        cwd=synced_setup,
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 25
    assert "SECURITY ERROR" in completed.stdout
    assert len(schedule_capture.read_text(encoding="utf-8").splitlines()) == 1
    assert not (local_app_data / "GiftCardRecon").exists()
