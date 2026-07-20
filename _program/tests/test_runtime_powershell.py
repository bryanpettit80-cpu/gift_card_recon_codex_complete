from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

import pytest


PROGRAM_ROOT = Path(__file__).resolve().parents[1]


@pytest.mark.skipif(sys.platform != "win32", reason="Windows junction integration test")
def test_runtime_reparse_guard_powershell() -> None:
    powershell = shutil.which("pwsh") or shutil.which("powershell.exe")
    assert powershell is not None, "PowerShell is required for the Windows runtime guard"

    fixture = PROGRAM_ROOT / "maintenance" / "test_runtime_reparse_guard.ps1"
    completed = subprocess.run(
        [
            powershell,
            "-NoLogo",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(fixture),
        ],
        cwd=PROGRAM_ROOT.parent,
        text=True,
        capture_output=True,
        timeout=60,
        check=False,
    )

    assert completed.returncode == 0, (
        f"PowerShell fixture failed with exit code {completed.returncode}\n"
        f"stdout:\n{completed.stdout}\n"
        f"stderr:\n{completed.stderr}"
    )
