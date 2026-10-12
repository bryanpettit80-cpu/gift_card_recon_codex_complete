# Reporting, corrections, and recovery

The reporting commands share the Decimal calculation and validation services. Weekly reporting compares Activity with Micros POS and tender evidence. Monthly close adds the fiscal Summary, Darden tie-out, and bound explanations. Each restaurant receives an independent result.

This guide describes the application interfaces. Installing or updating a particular workstation, auditing historical results, and changing operational folders are separate activities. Attachment intake, scheduling, notifications, and the Richmond Micros copier remain outside the reporting workflow.

## Normal operation

Use the existing Weekly and Monthly launchers after explicit setup. Reporting reuses the installed dependencies; it does not install or upgrade packages. The small local path configuration is written by `install.ps1`. `-ShowConfiguration` displays the configured roots without installing or creating business folders.

The PowerShell examples below run from the source repository root after setup. They use the saved operations and Dropbox paths. Supply `-OperationsRoot` and `-DropboxRoot` explicitly when using another workspace. Operator `.cmd` launchers live in the operations root and forward the same options.

Transaction types use an explicit allowlist covering activation, reload, their voids, redemption, redemption no NSF, and redemption void. Unsupported labels, Boolean money, nonfinite money, missing Activity folders, and ambiguous restaurant folders are actionable errors. Each restaurant receives its own result.

Activity and Summary roles are discovered from workbook content. Excel owner files are ignored. Summary dates come from explicit internal dates, captured intake metadata, or a legacy filename date. A newly renamed Summary with neither an internal date nor a matching intake record remains blocked: its period cannot safely be inferred from its totals.

Original captured inputs and archived evidence retain size/hash checks. Editable reports use relevant cell values, formulas, identities, and explanation bindings. Styles and ordinary Excel saves do not change business identity. Semantic Activity identity includes every normalized record and its multiplicity. It does not infer that two identical rows represent one transaction.

## Readiness and structured results

```powershell
.\_program\run_weekly.ps1 -Readiness
.\_program\run_monthly_close.ps1 -Readiness -Store 9355 -Period "FY27-M01"
```

Omit monthly `-Store` to assess both restaurants. The Python interface is also available through `python -m gift_card_recon readiness weekly` and `readiness monthly`, using an environment that can import the application, `--operations-root`, and the monthly `--period`. Readiness does not stage, extract, move, publish, or replay recovery records. Compressed Micros sources must be extracted through an execution/intake step before readiness can assess their contents. Pending recovery is reported as an action, not performed by a readiness request.

Execution supports JSON results, including restaurant, period, input/result/run identities, business disposition, artifacts, PDF status, correction details, and recovery actions. These interfaces use the same preparation and validation services as the manual commands. They do not connect a scheduler.

## Corrections and previous versions

Unchanged deliveries produce a duplicate/no-change result. A historical duplicate can use preserved evidence after the live Micros export has rotated beyond that week. Current partial or malformed evidence remains an error.

A changed input set first reports changed facts and a correction token. Review the facts, then pass that exact token once:

```powershell
.\_program\run_weekly.ps1 -Store 9355 -ApproveCorrection "9355=TOKEN"
.\_program\run_monthly_close.ps1 -Store 9355 -Period "FY27-M01" -ApproveCorrection "TOKEN"
```

Replace `TOKEN` with the exact proposal value. The token binds the previous result and proposed inputs. Recalculation rechecks that binding. The previous workbook, saved explanation, original evidence, and manifest are preserved before replacement. Repeated corrections create further versions. An explanation is reused only while its associated controls and amounts still match. A report saved while publication is in progress is retained and the replacement stops for review.

Weekly prior versions live beside the current weekly package in `YYYY-Www.revisions`. Monthly versions live under the store-period archive's `revisions` directory. Use recorded revision identities; the following values are placeholders:

```powershell
.\_program\run_weekly.ps1 -Store 9355 -RestoreRevision "9355=2026-W27/r0001-REVISIONID"
.\_program\run_monthly_close.ps1 -Store 9355 -Period "FY27-M01" -RestoreRevision "RESULT_ID"
```

The first request produces the required replacement token. Repeat the same command with its exact `-ApproveCorrection` value to preserve the current result and publish the selected evidence as another revision. Existing archives remain readable without an upfront migration. Restore verifies the current and selected reports and retained evidence; a damaged report or archive can require investigation before restore is allowed.

## PDF delivery

A validated monthly workbook can complete while PDF status is `pending`. PDF export errors do not discard the workbook or require another business correction. Retry separately:

```powershell
.\_program\run_monthly_close.ps1 -Store 9355 -Period "FY27-M01" -RetryPdf
```

The retry verifies the saved result before exporting and updates its delivery status. It stops if the workbook changes during export. Desktop Excel is required for this PDF exporter, not for workbook calculation. 7-Zip is needed only for `.7z` inputs.

## Interrupted publication

One local writer lock covers the workspace. Durable transaction records live under `_automation_runs/support`. Each published replacement records the prior copy and intended content before changing destinations. Replacement copies are verified and flushed; publication uses a rename on the destination volume. Temporary files beside an externally configured destination are recorded in the same journal.

The next execution restores an uncommitted publication or finishes committed cleanup. Changed source files are retained rather than deleted. Recovery preserves an unexpected destination edit and reports the journal requiring review. Readiness only reports this condition.

Completed workspaces can move with their inputs and archives. Interrupted journals rebase paths within the moved workspace; external destinations that cannot be remapped require recovery with the original configured roots. This is single-workstation coordination. Finish recovery and Dropbox synchronization before sequential handoff to another computer.

## Validation and rollback

`run_tests.ps1` uses an isolated test runtime and compact synthetic workspaces. The suite covers malformed inputs, independent stores, semantic repeats, bound corrections, explanations, missing PDF/retry, overlapping writers, locked/unavailable storage, and subprocess termination at publication boundaries. Private restaurant fixtures stay outside Git. Runtime fixtures cover another username/path, offline relocation, failed upgrades, and a terminated upgrade while preserving the previous active runtime.

No production archive migration is required to use existing results. To roll back the program, select the previous program version and its retained working runtime. New archive revisions and transaction records should be retained. Complete pending recovery with the new version before reverting to a version that cannot read its journal.

Implementation validation included five real Excel PDF exports on the development workstation and visual review of representative two-page reports. Cloud-provider outages and a physical second computer still require environment-specific acceptance checks; synthetic failures and same-machine relocation tests do not establish those outcomes. Process termination tests establish journal recovery at the tested boundaries, not resilience to every hardware or filesystem power-loss failure.

Legacy manifests without a semantic report baseline may still require the original workbook bytes for their first migration. A hash-verified historical archive establishes preservation; it does not independently recalculate the historical business result. Run a separately scoped audit when that assurance is needed.
