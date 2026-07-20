from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

import pytest


PROGRAM_ROOT = Path(__file__).resolve().parents[1]
POWERSHELL_EXECUTABLES = tuple(
    dict.fromkeys(
        executable
        for executable in (shutil.which("pwsh"), shutil.which("powershell.exe"))
        if executable is not None
    )
)


@pytest.mark.skipif(sys.platform != "win32", reason="Windows junction integration test")
@pytest.mark.parametrize(
    "powershell",
    POWERSHELL_EXECUTABLES or (None,),
    ids=lambda executable: Path(executable).name if executable else "missing",
)
def test_runtime_reparse_guard_powershell(powershell: str | None) -> None:
    if powershell is None:
        pytest.fail("PowerShell is required for the Windows runtime guard")

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
