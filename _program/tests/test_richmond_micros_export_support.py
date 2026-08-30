from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
from pathlib import Path

import pytest


pytestmark = pytest.mark.skipif(os.name != "nt", reason="Windows cmd.exe integration test")


REPO_ROOT = Path(__file__).resolve().parents[2]
SUPPORT_ROOT = REPO_ROOT / "_program" / "support" / "richmond_micros_export"
TRUSTED_INSTALLER_NAME = "Trusted-InstallDailyGiftCardCopyTask.ps1"
RELEASE_MANIFEST_NAME = "GiftCardExportReleaseManifest.json"
PAYLOAD_NAME = "Copy-GiftCardExportToDropbox.cmd"


def test_synced_installer_is_a_fail_closed_tombstone(tmp_path: Path) -> None:
    installer = tmp_path / "Install-DailyGiftCardCopyTask.cmd"
    shutil.copy2(SUPPORT_ROOT / installer.name, installer)

    completed = subprocess.run(
        [os.environ.get("COMSPEC", "cmd.exe"), "/d", "/c", installer.name],
        cwd=tmp_path,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 26, completed.stdout + completed.stderr
    assert "Do not install the Richmond gift-card task from the synced Dropbox folder" in (
        completed.stdout
    )

    installer_text = installer.read_text(encoding="utf-8").lower()
    assert "schtasks" not in installer_text
    assert "powershell" not in installer_text


def test_trusted_release_manifest_matches_the_committed_payload() -> None:
    payload_path = SUPPORT_ROOT / PAYLOAD_NAME
    manifest = json.loads((SUPPORT_ROOT / RELEASE_MANIFEST_NAME).read_text(encoding="utf-8"))

    assert manifest["schema_version"] == 1
    assert manifest["release_id"] == "richmond-gift-card-export-1"
    assert manifest["release_version"] == 1
    assert manifest["payload_filename"] == PAYLOAD_NAME
    assert manifest["payload_size_bytes"] == payload_path.stat().st_size
    assert manifest["payload_sha256"] == hashlib.sha256(payload_path.read_bytes()).hexdigest().upper()


def test_trusted_verifier_authenticates_before_staging_and_uses_system_scheduler() -> None:
    verifier_text = (SUPPORT_ROOT / TRUSTED_INSTALLER_NAME).read_text(encoding="utf-8")

    assert (
        "[Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)"
        in verifier_text
    )
    assert "$env:LOCALAPPDATA" not in verifier_text
    assert '$SchtasksPath = Join-Path $SystemDirectory "schtasks.exe"' in verifier_text
    assert '$CmdPath = Join-Path $SystemDirectory "cmd.exe"' in verifier_text
    assert "ReparsePoint" in verifier_text
    assert "Assert-RestrictedWriteAcl" in verifier_text
    assert "GetAccessControl" in verifier_text

    neutralize = verifier_text.index("& $SchtasksPath /Create /TN $TaskName /TR $SafeAction")
    source_check = verifier_text.index(
        'Assert-FingerprintMatchesManifest -Path $sourcePath -Manifest $manifest'
    )
    stage_copy = verifier_text.index(
        "Copy-Item -LiteralPath $sourcePath -Destination $stagePath"
    )
    stage_check = verifier_text.index(
        'Assert-FingerprintMatchesManifest -Path $stagePath -Manifest $manifest'
    )
    payload_copy = verifier_text.index(
        "Copy-Item -LiteralPath $stagePath -Destination $PayloadPath"
    )
    payload_check = verifier_text.index(
        'Assert-FingerprintMatchesManifest -Path $PayloadPath -Manifest $manifest'
    )
    activate = verifier_text.index(
        '& $SchtasksPath /Create /TN $TaskName /TR (\'"{0}"\' -f $PayloadPath)'
    )

    assert neutralize < source_check < stage_copy < stage_check < payload_copy < payload_check < activate
    assert "\nschtasks /" not in verifier_text.lower()


def test_trusted_verifier_has_valid_powershell_syntax() -> None:
    verifier_path = SUPPORT_ROOT / TRUSTED_INSTALLER_NAME
    command = (
        "$tokens = $null; "
        "$errors = $null; "
        "[System.Management.Automation.Language.Parser]::ParseFile("
        "$env:GIFT_CARD_TEST_VERIFIER, [ref]$tokens, [ref]$errors) | Out-Null; "
        "if ($errors.Count -gt 0) { $errors | ForEach-Object { Write-Error $_.Message }; exit 1 }"
    )
    env = os.environ.copy()
    env["GIFT_CARD_TEST_VERIFIER"] = str(verifier_path)
    completed = subprocess.run(
        [
            os.path.join(os.environ["SystemRoot"], "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
            "-NoLogo",
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            command,
        ],
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stdout + completed.stderr
