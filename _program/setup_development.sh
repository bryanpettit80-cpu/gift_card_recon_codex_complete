#!/usr/bin/env bash
set -euo pipefail

# Development only. Windows reporting uses install.ps1 and its external runtime.
repository_root="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)"
python_command="${PYTHON:-python3}"
venv_root="$repository_root/.venv"

"$python_command" -c 'import sys; sys.exit("Development requires Python 3.11 or newer.") if sys.version_info < (3, 11) else None'
"$python_command" -m venv "$venv_root"
"$venv_root/bin/python" -m pip install --require-hashes -r "$repository_root/_program/requirements.lock"
"$venv_root/bin/python" -m pip install --no-index --no-build-isolation --no-deps -e "$repository_root/_program[dev]"
"$venv_root/bin/python" -m pip check

printf 'Development environment ready: %s\n' "$venv_root/bin/python"
