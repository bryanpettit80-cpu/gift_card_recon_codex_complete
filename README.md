# Gift Card Reconciliation

Windows reporting application for Richmond (`9354`) and Virginia Beach (`9355`). It reconciles weekly gift-card Activity against Micros POS and tender evidence, carries completed activity into fiscal-month close, and checks the monthly Summary against the Darden credit memo.

The repository contains the application, tests, and operator-file templates. Restaurant inputs, reports, archives, and operational logs belong in the separate operations workspace.

## What is implemented

| Area | Current behavior |
| --- | --- |
| Input validation | Explicit transaction classifications, strict money parsing, restaurant/date/coverage checks, content-based report discovery, and independent results for each restaurant. |
| Portable setup | Explicit installation, saved local path configuration, reusable dependency environments, and separate reporting/test runtimes. |
| Publication and recovery | One local writer, durable recovery journals, verified previous copies, and recovery after interrupted publication. |
| Business identity | Unchanged records produce no-change results; harmless filenames, formatting, and ordinary Excel saves do not create a business correction. Original evidence retains byte/hash verification. |
| Corrections and delivery | Reviewed correction tokens, preserved revisions, selected-version restore, amount-bound explanations, and completed monthly workbooks with separately retryable PDF delivery. |
| Automation interface | Read-only readiness and structured execution results through the same validation services. Scheduling, attachment intake, and notifications remain separate. |

See [reporting, corrections, and recovery](docs/reporting-workflow.md) for commands and limits.

## First setup

Install Python 3.10 or newer on Windows. Download the program and make the required Dropbox inputs and archives available locally. From the repository root, run:

```powershell
.\_program\install.ps1 -OperationsRoot 'C:\Your Dropbox\Automations\Gift Card Reconciliation' -DropboxRoot 'C:\Your Dropbox'
```

Setup creates the local runtime and installs the operator launchers and guide in the operations workspace. Desktop Excel is needed for PDF export; 7-Zip is needed only when using `.7z` sources. Reporting commands reuse the installed environment and never install packages.

After setup, open the operations workspace and use its two launchers. Use `-ShowConfiguration` to inspect resolved paths or `-Readiness` to check available inputs before execution; see [readiness and structured results](docs/reporting-workflow.md#readiness-and-structured-results).

## Weekly reporting

1. Save one Monday-Sunday Gift Card Activity workbook in the correct store's `activity` inbox.
2. Close the workbook in Excel.
3. Click `Run Weekly Gift Card Reconciliation.cmd`.

The program retrieves that store's matching POS data, validates all seven business dates and tender evidence, creates the workbook, stages the Activity report for monthly close, and retains a hash-verified weekly evidence package. No POS control CSV entry is required.

Stores are processed independently. An empty inbox is a normal no-op. A malformed or incomplete report remains in its inbox, does not publish a workbook, and causes the launcher to identify the review item while still allowing the other store to run.

## Folders

The supported clean-root deployment keeps the whole operator workspace together
under `Dropbox\Automations` while leaving the external Micros exports at the
Dropbox root:

```text
Dropbox/
  Automations/
    Gift Card Reconciliation/  # complete operator workspace shown below
  micros_data/RC-Richmond-current/
  GETLinkedData-VB/
```

```text
00 START HERE - Gift Card Reconciliation.txt
Run Weekly Gift Card Reconciliation.cmd
Run Monthly Gift Card Close.cmd
01 Weekly Gift Card Activity Reports/
  9354 Richmond/activity/
  9355 Virginia Beach/activity/
02 Monthly Close Inputs/
  9354 Richmond/
  9355 Virginia Beach/
  Darden Reports - Drop Here/
03 Finished Reports/
  Weekly/
  Monthly Close/
  Monthly Close - Review Required/
04 Archive/
  Weekly Reconciliation/  # weekly source, POS evidence, workbook, and manifest
  Monthly Close/          # canonical close evidence and manifests
  Generated Reports/      # preserved historical workbooks
  Legacy Reconciliation/  # pre-current-process Darden material
  Cleanup Manifests/      # hash-verified organization records
_automation_runs/
Gift Card Reconciliation Automation/  # program-only Git repository
```

Operators normally use only the two launchers, weekly Activity inboxes, Darden inbox, and finished reports. `Gift Card Reconciliation Automation` contains the code and tests; `_automation_runs` contains operational logs, recovery records, and review quarantine. Test runtimes and fixtures stay outside the operator workspace.

Variance explanations are entered directly in the finished weekly workbook. Older companion forms under the fiscal period's `Variance Explanations` folder remain readable for compatibility.

The Python environment, package cache, compiled Python cache, and temporary extraction files are kept outside Dropbox under `%LOCALAPPDATA%\GiftCardRecon`. They are not part of the repository or monthly-close evidence.

## Automatic Weekly POS Controls

The normal weekly runner reads `DLYSYSTT.TXT` and `TENDER_DETAIL.TXT` from each store's configured external Micros export. It requires one correct-store Monday-Sunday Activity report, exact weekly POS coverage, valid money fields, and matching tender evidence. A scheduled Monday may be absent only when Activity, POS, and tender evidence for that day are all zero. Missing values are never replaced with Activity totals.

A completed store/week is retained under `04 Archive\Weekly Reconciliation` with the original Activity report, a compact seven-day POS/tender CSV, a preserved copy of the published workbook, and `weekly_manifest.json` containing sizes and SHA-256 hashes. Semantically unchanged reruns do not publish again. Changed records or controls produce a correction preview and require its exact replacement token; the previous result is retained as a numbered revision.

Whenever the absolute value of any weekly POS/tender control is greater than `$5.00`, open the finished weekly workbook's `Variance Explanation` tab, enter a brief plain-text explanation in the yellow `B15` cell, and save it in place. No separate explanation workbook is required. Monthly close validates relevant values, formulas, identity, and explanation bindings against the preserved baseline while allowing harmless formatting and Excel saves. Explained controls keep their monetary variance and receive `EXPLAINED` status; they can close with review.

Fiserv `Reload` transactions count toward gift-card issuance, and `Void Of Reload` reverses issuance. Both remain visible under their original labels in Raw Detail.

Completed weekly workbooks remain under `03 Finished Reports\Weekly`, including workbooks with a `REVIEW` status. The separate `Monthly Close - Review Required` folder is only for blocked monthly-close diagnostics.

## Workbook

The first tab is `Reconciliation`. It now includes:

- The POS tie-out rows
- A `Gift Card Activity File Totals` section with the activity file name, report period, row count, activations, redemptions, and net activity
- POS control totals

Other tabs keep the weekly, daily, raw detail, source file, and exception details.

## Runtime and relocation

Setup records the local paths in `%LOCALAPPDATA%\GiftCardRecon\paths.json`. Explicit runner parameters override those defaults. Standard nested and flat layouts work without a machine-specific launcher; an external program location uses the local configuration. Rerun setup after changing that external program location. Reporting never installs packages: a missing or changed dependency environment gives a clear setup instruction.

Setup installs the hash-verified `_program\requirements.lock` packages into a new permanent environment beneath `%LOCALAPPDATA%\GiftCardRecon\environments`. It validates that environment before atomically switching `active-runtime.json`; the previous environment and previous pointer remain available for recovery. Failed or interrupted upgrades leave the previous runtime selected. There is no in-place clearing of the reporting environment. Code is loaded from the invoking checkout, so relocating unchanged dependencies reuses the environment offline. Dependency identity uses `requirements.lock` and `requirements-bootstrap.in`, not checkout paths or workbook bytes. `_program\.venv` is no longer used.

Desktop Excel is required only for PDF export. Workbook calculations need Python and dependencies; 7-Zip is needed only for `.7z` sources. Complete Dropbox synchronization before handing work to another computer. No email credentials, machine trust provisioning, or reporting schedule are part of setup.

`pyproject.toml` remains the application dependency source of truth, `requirements-bootstrap.in` declares installer/build tooling, and `requirements.txt` remains a compatibility list. Maintainers must regenerate the cross-platform lock, including test and bootstrap dependencies, from `_program` after changing dependency declarations:

```powershell
uv pip compile pyproject.toml requirements-bootstrap.in --all-extras --universal --python-version 3.10 --generate-hashes --output-file requirements.lock
```

Temporary Micros extraction uses `%LOCALAPPDATA%\GiftCardRecon\temp\micros-extract`. Python bytecode, pytest state, and package downloads use `%LOCALAPPDATA%\GiftCardRecon\cache`.

## Test

From the repository root, run:

```powershell
.\_program\run_tests.ps1
```

The test runner uses a separate checkout-specific runtime and caches beneath `%LOCALAPPDATA%\GiftCardRecon\tests`. It cannot select the reporting runtime. Use `-SkipExcelIntegration` when desktop Excel integration is unavailable; `-RuntimeRoot` can select another isolated test location. Tests do not install operator assets or change the saved reporting paths.

Test fixtures use a unique short directory beneath `%LOCALAPPDATA%\GiftCardRecon\t` to avoid Windows path-length limits. `-TestTempRoot` changes its parent; pytest receives only a newly named child of that location.

## Readiness, corrections, and PDF retry

Both PowerShell runners accept `-ShowConfiguration` to display resolved local paths without creating operator folders. Weekly `-Readiness`, or monthly `-Readiness -Period FY27-M01`, returns JSON readiness without staging, extracting, publishing, or performing recovery. Both runners accept `-Json` for structured execution results. A correction requires the exact token shown by the assessment: weekly `-ApproveCorrection '9354=TOKEN'`, or monthly `-Store 9354 -Period FY27-M01 -ApproveCorrection TOKEN`. Monthly `-RetryPdf -Store 9354 -Period FY27-M01` retries PDF delivery from the completed workbook.

To select a retained report revision, use weekly `-RestoreRevision '9354=2026-W40/REVISION'` or monthly `-Store 9354 -Period FY27-M01 -RestoreRevision RESULT_ID`. Review the proposal, then repeat with its exact `-ApproveCorrection` token. The selected report is published as a new revision and intervening history remains preserved.

See [reporting, corrections, and recovery](docs/reporting-workflow.md) for the shared workflow, readiness limits, and sequential workstation handoff.

## Monthly Close From Micros

The Darden credit memo is the final checkbox for the month. Put every new Darden PDF in the shared inbox:

```text
02 Monthly Close Inputs\Darden Reports - Drop Here\
```

Then double-click `Run Monthly Gift Card Close.cmd`. A no-argument run scans every PDF in the inbox, reads the store and fiscal service period from the report, and processes Richmond and Virginia Beach independently. There is no background watcher.

The source folders remain store- and period-specific:

```text
02 Monthly Close Inputs\9355 Virginia Beach\FY27 M01 - Fiscal June\
  summary\
    07.05.2026 9355 Gift Card Summary.xlsx
  activity\
    five Monday-Sunday Gift Card Activity reports
  Variance Explanations\
    older companion forms, if created before embedded explanations
```

Weekly activity files are staged automatically by the weekly runner. Completed evidence is retained under `04 Archive\Monthly Close` with a SHA-256 close manifest.

### Close dispositions

- `CLOSED`: every required control passes.
- `CLOSED WITH REVIEW`: evidence is complete and remaining monetary differences are explained, or POS/tender differences remain within the existing `$5.00` review tolerance. This is a completed close.
- `REVIEW REQUIRED`: required explanations are missing or stale, or identity, completeness, coverage, archive integrity, or publication fails.

The Darden result is shown separately as `MATCHED` or `MISMATCHED`; it is not the overall close status. Summary-to-Activity and Darden-to-Summary differences require either an exact match or an explanation for that specific current control and amount. An explanation never changes the numerical result or waives source identity, date coverage, or archive checks.

For every week with a POS/tender control greater than `$5.00` in either direction, monthly close requires the matching weekly report's nonblank plain-text explanation. A period rollup is explained only when its contributing nonzero weekly controls have explanations. Older completed companion forms are still supported.

For remaining monthly differences, open the monthly review workbook's `Monthly Variance Explanations` tab, enter the reason beside the specific amount, save, close Excel, and rerun Monthly Close. The program validates the store, period, control and amount, retains the explanation source, and publishes `CLOSED WITH REVIEW` when no blocker remains. A changed amount requires a renewed explanation; monthly text cannot replace a missing required weekly explanation.

The runner requires exactly one correct-store activity report for every expected week, exact fiscal-date coverage, and both Micros evidence files. A scheduled Monday may be absent from Micros only when activity and tender evidence are also zero. An existing Monday POS row is included normally. Missing POS values are never replaced with activity totals.

### Deliverables

Successful close reports include a validated workbook and a matching PDF when export is available:

```text
03 Finished Reports\Monthly Close\<fiscal period>\
  Richmond_9354_<period>_Monthly_Close.xlsx
  Richmond_9354_<period>_Monthly_Close.pdf
  Virginia_Beach_9355_<period>_Monthly_Close.xlsx
  Virginia_Beach_9355_<period>_Monthly_Close.pdf
```

The first worksheet is an intentional two-page, letter-landscape executive accounting report:

- Page 1 shows the location and fiscal period, an overall status band, `Settlement Tie-Out` cards, a `Close Controls` table, and `Open Items Summary`.
- Page 2 repeats the location and period, then shows `Weekly Variance Detail`, marks explained variances, and includes `Variance Summary`, deduplicated `Review Items`, and `Evidence and Audit Trail`. The monthly workbook retains the full weekly narratives in `Variance Explanations` and monthly reasons beside their amounts in `Monthly Variance Explanations`, without expanding the fixed two-page PDF.

The report uses Arial, a navy/light-blue accounting palette, and green/amber/red only for assessed status. Settlement amounts use neutral accounting formatting, including negative values. Both formats carry report metadata plus generated-time, location, and page details in the footer. The fixed 85% print scale, merged-cell borders, and compact follow-up text keep the report readable at exactly two pages.

A validated workbook completes the close even when Excel PDF export fails. The manifest records PDF delivery as pending and the command reports the export error. A stale PDF is removed transactionally so it cannot appear to match the new workbook. Use `-RetryPdf` to export the saved result separately without recalculating or approving another business correction.

A blocked run always attempts a red diagnostic workbook and an Excel-exported diagnostic PDF under `03 Finished Reports\Monthly Close - Review Required`. If Excel cannot create the PDF, the new workbook may be published by itself; the command reports the original close blocker, the exact PDF export error, the authoritative workbook path, and that no diagnostic PDF was published for that run. Any older same-named diagnostic PDF is retired transactionally first, so an old PDF cannot appear to match the new workbook. If a locked old diagnostic prevents that retirement, the old pair is preserved and the new diagnostic is not published; the command reports both the original blocker and the diagnostic-publication failure. A locked canonical file likewise fails clearly and never causes an alternate filename.

### Archive-backed reissues

Use the dedicated archive mode to reproduce a completed store-period from its retained evidence instead of manually assembling source paths:

```powershell
& ".\Run Monthly Gift Card Close.cmd" `
  -Store 9355 `
  -Period FY27-M01 `
  -ReissueFromArchive
```

`-ReissueFromArchive` requires both `-Store` and `-Period`. It cannot be combined with `-InputDir`, `-DardenPath`, or `-MicrosPath`; the program derives the canonical archived Summary, activity, Darden, and Micros paths from the existing close manifest. It verifies every retained source against that manifest, rejects missing, changed, or out-of-archive evidence, and permits the archived Micros snapshot only after containment validation.

Archive mode forces the equivalents of `-NoStageWeekly` and `-NoCleanup`. It does not stage weekly files, delete sources, or touch live input folders. Reissued canonical artifacts still follow the normal transactional workbook/PDF publication rules.

### Explicit reruns

Store, period, Darden, input, output, and Micros options remain available for controlled reruns:

```powershell
& ".\Run Monthly Gift Card Close.cmd" `
  -Store 9355 `
  -Period FY27-M01 `
  -DardenPath "C:\path\to\Darden credit memo.pdf" `
  -InputDir ".\04 Archive\Monthly Close\9355\FY27 M01 - Fiscal June" `
  -NoStageWeekly `
  -NoCleanup
```

`-Period 2026-06` maps to Darden Fiscal June 2026 (`FY27-M01`), covering `2026-06-01` through `2026-07-05`.

All new evidence is written beneath `04 Archive\Monthly Close`. A historical lowercase `monthly-close` path remains readable when supplied explicitly with `-InputDir`, but it is never used as a silent fallback.

Default Micros sources are location-specific:

- Richmond / `9354`: the external Dropbox folder `micros_data\RC-Richmond-current`
- Virginia Beach / `9355`: the external Dropbox folder `GETLinkedData-VB`

Those sources remain at the Dropbox root when the operator workspace moves to
`Dropbox\Automations\Gift Card Reconciliation`. The PowerShell entrypoints
resolve the standard legacy and clean-root layouts and pass an explicit Dropbox
root to Python. For a nonstandard layout, supply `-DropboxRoot` to either
entrypoint. Relative operator input/output overrides still resolve from
`-OperationsRoot`.

`-MicrosPath` may point to the configured location source or to a store-identified archived snapshot (an extracted folder, `.zip`, or `.7z`). Arbitrary folders and the other location's source are rejected.

## Program-Only Repository and Operator Assets

The GitHub repository is the source history. A development checkout may live outside Dropbox; the standard deployed copy is `Gift Card Reconciliation Automation` inside the operations workspace. Both contain program files only. A source update does not establish that an operator's deployed copy has been updated: deployment should verify the intended commit and copied files separately.

Both PowerShell entrypoints accept `-OperationsRoot` and `-DropboxRoot`. The parent-facing launchers pass their own folder explicitly, so inputs, reports, archives, logs, and review files remain outside the nested Git repository. Relative override paths are resolved from the operations root, not from the code checkout. Store Micros exports resolve separately from the Dropbox root and remain outside `Automations`.

After a clean checkout or operator-file refresh, deploy and SHA-256-verify the guide, launchers, drop-folder notes, and required folders with:

```powershell
.\_program\install_operator_assets.ps1 -OperationsRoot "C:\Users\bryan\Dropbox\Automations\Gift Card Reconciliation"
```

See `docs/WORKSPACE_RELOCATION_RUNBOOK.md` before relocating an existing live
workspace. Historical absolute paths in weekly manifests are provenance only;
duplicate validation discovers the current report by its content and continues to
verify contained original evidence by size and SHA-256. Editable reports are
checked by relevant values, formulas, identity, and explanation bindings.

## Source map and documentation

| Path | Purpose |
| --- | --- |
| [`_program/src/gift_card_recon`](./_program/src/gift_card_recon) | Parsers, calculations, weekly/monthly services, readiness, publication, and report generation. |
| [`_program/tests`](./_program/tests) | Synthetic business, publication/recovery, runtime, and Excel integration tests. |
| [`_program/install.ps1`](./_program/install.ps1) | Explicit workstation setup and path configuration. |
| [`_program/run_weekly.ps1`](./_program/run_weekly.ps1), [`_program/run_monthly_close.ps1`](./_program/run_monthly_close.ps1) | Supported Windows reporting entrypoints. |
| [`templates`](./templates) | Canonical operator launchers, start guide, and inbox notes. |
| [`docs/reporting-workflow.md`](./docs/reporting-workflow.md) | Readiness, corrections, restore, PDF retry, and recovery. |
| [`docs/WORKSPACE_RELOCATION_RUNBOOK.md`](./docs/WORKSPACE_RELOCATION_RUNBOOK.md) | Controlled relocation of an existing operations workspace. |
| [`docs/NUMBERED_DROPBOX_MIGRATION_RUNBOOK.md`](./docs/NUMBERED_DROPBOX_MIGRATION_RUNBOOK.md) | Historical-layout migration procedure; ordinary reporting does not require migration. |

## Validation scope

The [CI workflow](./.github/workflows/ci.yml) tests Python on Linux and Windows; real desktop Excel integration is a separate workstation check. Local validation has exercised five real Excel PDF exports and representative two-page layouts. Synthetic interruption and relocation tests cover the tested boundaries. Physical second-computer operation, cloud-provider outages, and arbitrary hardware power loss require their own environment-specific acceptance checks.

Use the test results for the relevant commit and the deployment record for the relevant workstation. Historical report preservation and hash checks do not establish an independent recalculation of historical results.
