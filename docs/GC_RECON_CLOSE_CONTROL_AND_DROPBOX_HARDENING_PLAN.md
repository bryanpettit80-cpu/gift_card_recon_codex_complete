# GC Recon Close-Control and Dropbox Hardening Plan

Approved for implementation on July 11, 2026.

> July 16 organization addendum: the operator workspace uses numbered folders and a deployed program snapshot with no `.git` metadata. GitHub is the authoritative source, and a local clone is the sole Git working copy. Current live paths are `02 Monthly Close Inputs`, `03 Finished Reports`, and `04 Archive`; the unnumbered paths below describe the pre-migration implementation snapshot preserved by this plan.

- Remote implementation branch: `agent/gc-recon-close-hardening`
- Draft pull request: [#5 — Harden GC Recon monthly close and Dropbox workflow](https://github.com/bryanpettit80-cpu/gift_card_recon_codex_complete/pull/5)

## Validated Baseline Before Final Polish and Reissue

The results below are the July 11 pre-polish baseline. They define the accounting outcomes that the archive-backed reissue must preserve; they are not a claim that the revised reports, final reissue, or GitHub delivery have already been verified.

- Richmond, store 9354: **CLOSED WITH REVIEW**. Period POS net variance is `($2.43)` and the largest weekly absolute variance is `$2.44`.
- Virginia Beach, store 9355: **CLOSED**.
- Darden matches the Summary to the cent for both locations.
- Prior validation: 107 tests passed; 8 legacy or optional-fixture tests were intentionally skipped.
- Dropbox consolidation completed and was rerun idempotently: all 124 recorded operations were already complete on the verification run.
- Generated close reports and accounting evidence remain in Dropbox and are intentionally excluded from GitHub.

## Summary

- Make `RICHMOND — STORE 9354` and `VIRGINIA BEACH — STORE 9355` the dominant report headings.
- Replace the Darden-only greenlight with one centralized assessment:
  - `CLOSED`: every required control passes.
  - `CLOSED WITH REVIEW`: evidence is complete, Darden matches, and all nonzero POS/tender variances are no more than `$5.00` at both weekly and period levels.
  - `REVIEW REQUIRED`: identity, completeness, coverage, Darden, archive, or larger-variance control fails.
- Reissue June FY27 reports as:
  - Richmond 9354: **CLOSED WITH REVIEW** for the `$2.43` period POS variance.
  - Virginia Beach 9355: **CLOSED**.
- Present both reports as restrained executive accounting certificates, using a consistent two-page layout and status colors only for assessed outcomes.

## Close Controls and Operator Flow

- Introduce central `StoreConfig`, `ControlOutcome`, and `CloseAssessment` types. Rename the Darden-only `closed` property to `darden_matched`; renderers must consume the assessment without recalculating status.
- Validate all evidence before closing:
  - Summary must contain exactly one exact store row; remove wrong-store fallback and strictly parse required money.
  - Activity reports must identify the correct store, provide exactly one Monday–Sunday report per expected week, and contain no duplicate, overlapping, missing, or out-of-period data.
  - Never substitute activity totals for missing POS totals.
  - Require exact Micros coverage except scheduled Mondays. A missing Monday is accepted only when activity and tender evidence are also zero; existing Monday POS data is included normally.
  - Require and normalize `TENDER_DETAIL.TXT`; missing or malformed tender evidence is blocking.
  - Validate the configured Micros source and column layout separately for each location.
  - Require exact-cent Darden and Summary-to-activity matches. Apply the `$5.00` amber threshold independently to every weekly and period POS/tender control.
- Replace stale launcher defaults with a shared Dropbox inbox:
  - `Monthly Close\Darden Reports - Drop Here`
  - A no-argument run scans all PDFs, derives store and fiscal period from content, and processes both locations independently.
  - Preserve explicit store, period, and file-path options for reruns.
  - Do not add a background watcher; the double-click runner initiates the scan.
- Block stale-output mistakes: a locked canonical output must fail clearly rather than create an alternate filename, and failed runs must not point operators toward an older workbook.

## Reports, Evidence, and Archiving

- Generate matching workbook and PDF artifacts under:
  - `Output\Monthly Close\<fiscal period>\Richmond_9354_<period>_Monthly_Close.*`
  - `Output\Monthly Close\<fiscal period>\Virginia_Beach_9355_<period>_Monthly_Close.*`
- Successful closes require a verified workbook/PDF pair. A PDF-export failure prevents canonical publication, archiving, and source cleanup.
- Blocked runs always attempt a red diagnostic workbook and Excel-exported PDF under `Output\Review Required`; they do not archive inputs or publish canonical close reports.
- If diagnostic PDF export fails, publish the workbook alone only after transactionally retiring any older same-named PDF. Report the original close blocker, the exact PDF error, the authoritative workbook path, and that no diagnostic PDF was published.
- If a locked older diagnostic PDF cannot be retired, preserve the older pair and report the diagnostic-publication failure without masking the original close blocker.
- Use an intentional two-page report:
  - Page 1: location/fiscal-period heading, overall status band, `Settlement Tie-Out` cards, `Close Controls`, and `Open Items Summary`.
  - Page 2: repeated location/period heading, `Weekly Variance Detail`, `Variance Summary`, grouped/deduplicated `Review Items`, and `Evidence and Audit Trail`.
- Format in letter landscape with Arial: 19-point title, 10-point subtitle/body, 11-point section headers, and 9.5-point secondary notes. Use a fixed 85% print scale and condense prose instead of shrinking the page further.
- Use navy `17365D`, light blue `D9EAF7`, and neutral white/light-gray rows. Reserve green, amber, and red for assessed status; do not color negative settlement amounts red unless their control is actually in review or blocked.
- Apply borders and alignment across every cell in merged ranges. Use ASCII hyphens, reader-friendly labels, and meaningful workbook/PDF title, subject, creator, description, generated-time, location, and page metadata.
- Distinguish `Darden: MATCHED` from the overall close status. Use green only for fully clean closes and amber for `CLOSED WITH REVIEW`.
- Show both largest weekly variance and period-net variance, use accounting currency formatting, and replace internal source-folder names with reader-friendly labels.
- Build follow-up items from every control outcome; passed weekly rows use `-`, zero-only reports show `$0.00 - No weekly variance`, and “No exceptions” appears only when no review or blocking item exists.
- Export PDF from the canonical workbook through Windows Excel automation so the two formats remain identical. Validate that the PDF opens, contains the location heading, and has exactly two pages; export failure blocks canonical publication and archiving while using the diagnostic-only fallback above.
- Snapshot and hash the Summary, activity files, Darden PDF, `DLYSYSTT.TXT`, and `TENDER_DETAIL.TXT`. Record canonical archive-relative paths and a close manifest.
- Make closeout transactional: render and verify temporary artifacts, copy and hash-verify all evidence, atomically publish the workbook/PDF, then remove live inputs and prune empty period folders. Partial failures leave original evidence intact.

## Archive-Backed Reissue and Delivery

- Add `-ReissueFromArchive` to the PowerShell launcher and `--reissue-from-archive` to the CLI. Require store and fiscal period; reject manual input-directory, Darden, or Micros overrides.
- Derive the archived Summary, activity, Darden, and Micros paths from the canonical store-period archive. Verify each source against the existing close manifest and reject missing, changed, or out-of-archive evidence before rendering.
- Permit the manifest-selected archived Micros snapshot only after archive-containment validation. Force no weekly staging and no source cleanup so a reissue cannot alter live or retained evidence.
- Before reissuing, copy each current workbook, PDF, and close manifest to `Archive - Old Files\Generated Reports\Monthly Superseded\<store>\FY27-M01\20260711-original-format\` and record original paths, sizes, and SHA-256 hashes in Dropbox-only `snapshot_manifest.json`.
- Reissue Richmond and Virginia Beach as one release. Preserve the baseline dispositions above; if either reissue or visual check fails, restore both original report pairs and manifests.
- Update only the tracked consolidation postflight hashes/sizes affected by the new manifests, then require its read-only dry run to verify all 124 operations remain complete.
- Merge current `origin/main` into the hardening branch without rebasing or force-pushing. Push the intended tracked changes to PR #5, require all three CI jobs, mark the PR ready, and squash-merge.
- After merge, verify exact local/remote tree equality before repointing local `main` and deleting the local feature branch. Confirm GitHub `main`, the local checkout, and Dropbox artifacts/manifests are synchronized.

## Code and Dropbox Organization

- Split the monolithic report and monthly-close files into focused assessment, report-rendering, PDF-export, and archive components. Write all workbook sections in one pass instead of saving, reopening, and appending.
- Pass `generated_at` into rendering for reproducible tests and centralize location names, thresholds, closed weekdays, source labels, and Micros configuration.
- Move the virtual environment, extraction workspace, and caches to `%LOCALAPPDATA%\GiftCardRecon`; install only when missing or when the dependency specification changes. Remove the validated legacy `_program\.venv` afterward.
- Add `.gitattributes` to eliminate line-ending churn.
- Perform safe Dropbox consolidation with a SHA-256 cleanup manifest:
  - Move the 11-file legacy Darden reconciliation folder intact beneath the canonical legacy archive and leave a redirect at its old location.
  - Preserve the six unique legacy weekly workbooks in the generated-report archive.
  - Remove only the 30 files already proven to have identical copies elsewhere.
  - Remove empty completed-June inbox shells and stale caches/temp data.
  - Migrate `Archive - Old Files\monthly-close` to the title-cased canonical structure while retaining backward-compatible reads.
- Keep generated reports, close manifests, source evidence, snapshots, POS controls, caches, and temporary accounting data out of Git; deliver only intended tracked code, tests, documentation, and consolidation-check changes through PR #5.

## Test and Acceptance Plan

- Add negative tests for wrong-store Summary/activity/Darden, duplicate and overlapping weeks, malformed money, period mismatches, missing tender evidence, wrong Micros source, partial coverage, output locking, PDF failure, and archive failure.
- Test scheduled Monday handling for both locations, including a Monday with real activity/POS data.
- Test all status outcomes, including Richmond’s amber result and Virginia Beach’s green result.
- Verify review reports cannot be mistaken for canonical close artifacts.
- Test workbook-only diagnostics, exact PDF-error reporting, stale-PDF retirement, locked stale artifacts, rollback, and nonzero CLI/wrapper exits.
- Test archive-reissue manifest verification, archive containment, required store/period, conflicting-option rejection, and forced no-staging/no-cleanup behavior.
- Verify source hashes, archive paths, cleanup idempotency, shared-inbox routing, multi-store runs, and nonzero wrapper exit codes.
- Keep Linux CI green with PDF automation mocked; run Windows/Python 3.14 unit and wrapper smoke checks in CI, while retaining the real Windows Excel/PDF integration test as a local workstation check.
- Reconcile the real June sources independently, regenerate both artifact pairs, verify workbook/PDF numerical parity and manifest hashes, render all four PDF pages, and visually inspect readability, borders, spacing, clipping, and exact two-page output.
- Require `git diff --check`, a clean tracked worktree, consolidation dry-run success, and passing GitHub CI on Linux with Python 3.10, 3.11, and 3.12 plus the Windows/Python 3.14 smoke job before merge.

## Assumptions

- Both locations use Monday as their recurring scheduled closed weekday, subject to evidence-based open-day detection.
- The amber limit is `$5.00` absolute for each weekly control and its period aggregate; any larger value blocks close.
- March 2026 material remains archived historical evidence but is not used in current logic or report narrative.
- Workbook and PDF are both required canonical monthly deliverables. A blocked diagnostic may be workbook-only only when Excel PDF export fails and the operator receives the explicit, stale-safe notification defined above.

## Numbered Operations Layout Addendum

- Parent launchers pass `-OperationsRoot` to the deployed `Gift Card Reconciliation Automation` snapshot. All relative input, report, archive, and review paths resolve from that explicit operations root.
- Weekly Activity inboxes are `01 Weekly Gift Card Activity Reports\9354 Richmond\activity` and `01 Weekly Gift Card Activity Reports\9355 Virginia Beach\activity`; monthly inputs use the corresponding labeled store folders under `02 Monthly Close Inputs`.
- Finished weekly reports, monthly close pairs, and diagnostics live under `03 Finished Reports`; retained evidence lives under `04 Archive`; logs, QA, quarantine, and test output live under `_automation_runs`.
- External Micros sources remain Dropbox siblings of the operations root. The local Git repository contains program code, tests, documentation, and versioned operator-asset templates; the Dropbox program snapshot contains only deployed runtime files and a hash manifest.
- `_program\install_operator_assets.ps1` creates the required operator folders and copies the START HERE guide, both parent launchers, and drop-folder notes from tracked templates, verifying every deployed file by SHA-256.
