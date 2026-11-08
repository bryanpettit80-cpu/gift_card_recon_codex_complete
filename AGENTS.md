# Codex Notes

This repository contains only the Gift Card Reconciliation automation program. Do not commit Activity reports, Summary workbooks, Darden PDFs, POS evidence, generated reports, manifests, archives, logs, backups, screenshots, or other restaurant data.

For the nested Windows deployment, the live Dropbox operations folder is the
parent of the program repository. In the clean-root layout it is:

`C:\Users\bryan\Dropbox\Automations\Gift Card Reconciliation`

The operator entry points live outside the repository:

- `Run Weekly Gift Card Reconciliation.cmd`
- `Run Monthly Gift Card Close.cmd`

Both parent launchers must pass `-OperationsRoot` to the PowerShell runner. Inputs, reports, archives, and `_automation_runs` belong under the parent operations folder; the Python runtime and caches belong under `%LOCALAPPDATA%\GiftCardRecon`. The Richmond and Virginia Beach Micros export folders remain at the Dropbox root, outside `Automations`, and must not be moved or committed. The runners infer that root for the standard legacy and `Automations` layouts and accept `-DropboxRoot` as an explicit override.

Keep the files under `templates` as the canonical Dropbox-facing operator assets. Refresh the parent guide, launchers, and drop-folder notes with `_program\install_operator_assets.ps1`; the installer must verify every deployed file by SHA-256.

For program, dependency, or runtime changes on Windows, run `_program\run_tests.ps1` before delivery. Use `-SkipExcelIntegration` when desktop Excel is unavailable or the change does not affect PDF layout/export; report that exclusion. Tests use a separate local runtime and temporary fixtures, not the reporting runtime.

On Linux or in Codex Cloud, run `bash _program/setup_development.sh` from the repository root, then run the synthetic suite from `_program`:

```bash
cd _program
../.venv/bin/python -m pytest -q \
  --ignore=tests/test_actual_files_if_present.py \
  --ignore=tests/test_windows_pdf_integration.py
```

Do not run Windows installers, migrations, scheduled-task setup, or live reporting in Cloud. Linux validation does not verify desktop Excel COM, Windows launcher/runtime behavior, Dropbox synchronization, or workstation credentials. Use only synthetic fixtures there; do not copy restaurant evidence into a development environment.

Test migrations only with `_program\maintenance\migrate_to_numbered_layout.ps1` in dry-run mode or with its isolated fixture script unless the user has approved the exact migration fingerprint.

## Repository map and pull requests

- `_program/src/gift_card_recon/` contains parsing, reconciliation, publication, and reporting; `_program/tests/` contains regression fixtures. Keep money calculations and validation separate from workbook/PDF presentation.
- Use [README.md](README.md) for operator behavior and [docs/reporting-workflow.md](docs/reporting-workflow.md) for reporting changes. Use the migration and relocation runbooks in `docs/` only when changing those workflows.
- Follow [.github/pull_request_template.md](.github/pull_request_template.md). Record the actual platform, commands, and exclusions; Linux or hosted Windows success does not establish desktop Excel rendering or live Dropbox behavior.
- For documentation-only changes, check referenced paths, commands, and the diff. Do not run reporting or deployment solely to validate documentation.
- Use a task branch and a draft PR. Merge only when requested, after required checks and reviews pass; this file grants no standing merge permission.

## Code Review Rules

- Preserve signed activation/reload, redemption, and void classifications and decimal arithmetic. Unknown transaction labels require review; do not hide discrepancies by changing a sign or replacing a missing control with zero.
- Preserve store, period, complete business-date coverage, correction-token, and hash-bound evidence checks. An unchanged rerun must remain idempotent; corrections must retain prior published evidence.
- Keep untrusted text literal in Excel and preserve the separation between user explanations and calculated controls. An explanation can change review status but must not erase the monetary variance.
