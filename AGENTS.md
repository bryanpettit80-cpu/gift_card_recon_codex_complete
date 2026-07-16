# Codex Notes

This repository contains only the Gift Card Reconciliation automation program. Do not commit Activity reports, Summary workbooks, Darden PDFs, POS evidence, generated reports, manifests, archives, logs, backups, screenshots, or other restaurant data.

GitHub is the authoritative source and history. Use a local clone as the sole Git working copy for development, tests, commits, and pushes. A profile-relative example is:

```powershell
$repoRoot = Join-Path $env:USERPROFILE "Documents\Repos\gift_card_recon_codex_complete"
```

The live Dropbox operations folder is separate from the Git checkout:

```powershell
$operationsRoot = Join-Path $env:USERPROFILE "Dropbox\Gift Card Reconciliation"
```

Dropbox contains restaurant evidence and a deployed `Gift Card Reconciliation Automation` program snapshot. The deployed snapshot must not contain `.git`, virtual environments, caches, tests, or other development state. Deploy only through `_program\maintenance\deploy_operator_program.ps1`, which writes the hash inventory to `Gift Card Reconciliation Automation\deployment-manifest.json`.

The operator entry points live outside the repository:

- `Run Weekly Gift Card Reconciliation.cmd`
- `Run Monthly Gift Card Close.cmd`
- `Check Gift Card Reconciliation Health.cmd`

Both parent launchers must pass `-OperationsRoot` to the PowerShell runner. Inputs, reports, archives, and `_automation_runs` belong under the parent operations folder; the Python runtime and caches belong under `%LOCALAPPDATA%\GiftCardRecon`. The Richmond and Virginia Beach Micros export folders are external Dropbox siblings and must not be moved or committed.

Keep the files under `templates` as the canonical Dropbox-facing operator assets. Refresh the guide, launchers, and drop-folder notes with `_program\install_operator_assets.ps1`; the installer must verify every deployed file by SHA-256. The health launcher calls `_program\check_operator_health.ps1` from the deployed snapshot. It may use a temporary output write probe and an explicitly requested JSON report, but it must not alter accounting inputs or publish reconciliation reports.

Run `_program\run_tests.ps1` from the local clone before delivery. Test migrations only with `_program\maintenance\migrate_to_numbered_layout.ps1` in dry-run mode or with its isolated fixture script unless the user has approved the exact migration fingerprint.
