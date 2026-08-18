# Codex Notes

This repository contains only the Gift Card Reconciliation automation program. Do not commit Activity reports, Summary workbooks, Darden PDFs, POS evidence, generated reports, manifests, archives, logs, backups, screenshots, or other restaurant data.

The live Dropbox operations folder is the parent of this nested repository. In
the clean-root layout it is:

`C:\Users\bryan\Dropbox\Automations\Gift Card Reconciliation`

The operator entry points live outside the repository:

- `Run Weekly Gift Card Reconciliation.cmd`
- `Run Monthly Gift Card Close.cmd`

Both parent launchers must pass `-OperationsRoot` to the PowerShell runner. Inputs, reports, archives, and `_automation_runs` belong under the parent operations folder; the Python runtime and caches belong under `%LOCALAPPDATA%\GiftCardRecon`. The Richmond and Virginia Beach Micros export folders remain at the Dropbox root, outside `Automations`, and must not be moved or committed. The runners infer that root for the standard legacy and `Automations` layouts and accept `-DropboxRoot` as an explicit override.

Keep the files under `templates` as the canonical Dropbox-facing operator assets. Refresh the parent guide, launchers, and drop-folder notes with `_program\install_operator_assets.ps1`; the installer must verify every deployed file by SHA-256.

Run `_program\run_tests.ps1` before delivery. Test migrations only with `_program\maintenance\migrate_to_numbered_layout.ps1` in dry-run mode or with its isolated fixture script unless the user has approved the exact migration fingerprint.
