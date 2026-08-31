from __future__ import annotations

import hashlib
import os
import re
import shutil
import subprocess
from pathlib import Path

import pytest


pytestmark = pytest.mark.skipif(os.name != "nt", reason="Windows cmd.exe integration test")


REPO_ROOT = Path(__file__).resolve().parents[2]
SUPPORT_ROOT = REPO_ROOT / "_program" / "support" / "richmond_micros_export"
TRUSTED_INSTALLER_NAME = "Trusted-InstallDailyGiftCardCopyTask.ps1"
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


def test_trusted_verifier_embeds_the_committed_payload_fingerprint() -> None:
    payload_path = SUPPORT_ROOT / PAYLOAD_NAME
    verifier_text = (SUPPORT_ROOT / TRUSTED_INSTALLER_NAME).read_text(encoding="utf-8")
    sha_match = re.search(r'^\$TrustedPayloadSha256 = "([A-F0-9]{64})"$', verifier_text, re.MULTILINE)
    size_match = re.search(r"^\$TrustedPayloadSizeBytes = \[int64\](\d+)$", verifier_text, re.MULTILINE)

    assert not list(SUPPORT_ROOT.glob("*Manifest*.json"))
    assert "GiftCardExportReleaseManifest.json" not in verifier_text
    assert sha_match is not None
    assert size_match is not None
    assert int(size_match.group(1)) == payload_path.stat().st_size
    assert sha_match.group(1) == hashlib.sha256(payload_path.read_bytes()).hexdigest().upper()


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
        'Assert-FingerprintMatchesRelease -Path $sourcePath -Release $release'
    )
    stage_copy = verifier_text.index(
        "Copy-Item -LiteralPath $sourcePath -Destination $stagePath"
    )
    stage_check = verifier_text.index(
        'Assert-FingerprintMatchesRelease -Path $stagePath -Release $release'
    )
    payload_copy = verifier_text.index(
        "Copy-Item -LiteralPath $stagePath -Destination $PayloadPath"
    )
    payload_check = verifier_text.index(
        'Assert-FingerprintMatchesRelease -Path $PayloadPath -Release $release'
    )
    activate = verifier_text.index(
        '& $SchtasksPath /Create /TN $TaskName /TR (\'"{0}"\' -f $PayloadPath)'
    )

    assert neutralize < source_check < stage_copy < stage_check < payload_copy < payload_check < activate
    assert "\nschtasks /" not in verifier_text.lower()


def test_trusted_verifier_compares_equivalent_windows_paths_safely() -> None:
    verifier_path = SUPPORT_ROOT / TRUSTED_INSTALLER_NAME
    command = (
        "$tokens = $null; "
        "$errors = $null; "
        "$ast = [System.Management.Automation.Language.Parser]::ParseFile("
        "$env:GIFT_CARD_TEST_VERIFIER, [ref]$tokens, [ref]$errors); "
        "$function = $ast.Find({ param($node) "
        "$node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and "
        "$node.Name -eq 'Test-EquivalentWindowsPath' }, $true); "
        "if ($null -eq $function) { Write-Error 'Path comparison function is missing.'; exit 1 }; "
        "Invoke-Expression $function.Extent.Text; "
        "$left = Join-Path ([System.IO.Path]::GetTempPath()) 'RichmondMicrosExport'; "
        "$right = $left.ToUpperInvariant() + [System.IO.Path]::DirectorySeparatorChar; "
        "if (-not (Test-EquivalentWindowsPath -Left $left -Right $right)) { "
        "Write-Error 'Equivalent Windows paths were rejected.'; exit 1 }; "
        "$other = Join-Path ([System.IO.Path]::GetTempPath()) 'OtherExport'; "
        "if (Test-EquivalentWindowsPath -Left $left -Right $other) { "
        "Write-Error 'Different Windows paths were accepted.'; exit 1 }"
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
