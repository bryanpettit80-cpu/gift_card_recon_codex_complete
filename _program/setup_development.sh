#!/usr/bin/env bash
set -euo pipefail

# Development only. Windows reporting uses install.ps1 and its external runtime.
repository_root="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd -P)"
python_command="${PYTHON:-python3}"
venv_root="$repository_root/.venv"
ownership_marker="$venv_root/.gift-card-recon-development"
expected_marker="gift-card-recon-development-v1
$repository_root"

if [[ -e "$venv_root" || -L "$venv_root" ]]; then
    if [[ -L "$venv_root" || ! -d "$venv_root" || -L "$ownership_marker" || ! -f "$ownership_marker" ]] ||
       [[ "$(cat -- "$ownership_marker")" != "$expected_marker" ]]; then
        printf 'Refusing to replace an unrecognized development environment: %s\n' "$venv_root" >&2
        exit 1
    fi
fi

# Resolve the base interpreter before clearing a managed environment, including
# when PYTHON or an activated PATH points at that environment's own interpreter.
base_python="$("$python_command" -c 'import os, sys; sys.exit("Development requires Python 3.11 or newer.") if sys.version_info < (3, 11) else None; print(os.path.realpath(getattr(sys, "_base_executable", sys.executable)))')"
if [[ ! -x "$base_python" || "$base_python" == "$venv_root"/* ]]; then
    printf 'A base Python interpreter outside %s is required.\n' "$venv_root" >&2
    exit 1
fi
"$base_python" -m venv --clear "$venv_root"
printf '%s\n' "$expected_marker" > "$ownership_marker"
"$venv_root/bin/python" -m pip install --require-hashes -r "$repository_root/_program/requirements.lock"
"$venv_root/bin/python" -m pip install --no-index --no-build-isolation --no-deps -e "$repository_root/_program[dev]"
"$venv_root/bin/python" -m pip check

printf 'Development environment ready: %s\n' "$venv_root/bin/python"
