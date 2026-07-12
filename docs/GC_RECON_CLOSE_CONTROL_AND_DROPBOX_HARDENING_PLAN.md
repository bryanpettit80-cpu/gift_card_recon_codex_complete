# GC Recon Close-Control and Dropbox Hardening Plan

Approved for implementation on July 11, 2026.

- Remote implementation branch: `agent/gc-recon-close-hardening`
- Draft pull request: [#5 — Harden GC Recon monthly close and Dropbox workflow](https://github.com/bryanpettit80-cpu/gift_card_recon_codex_complete/pull/5)

## Implementation Status

- Richmond, store 9354: **CLOSED WITH REVIEW**. Period POS net variance is `($2.43)` and the largest weekly absolute variance is `$2.44`.
- Virginia Beach, store 9355: **CLOSED**.
- Darden matches the Summary to the cent for both locations.
- Validation: 107 tests passed; 8 legacy or optional-fixture tests were intentionally skipped.
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
- Blocked runs create red diagnostic workbook and PDF files under `Output\Review Required`; they do not archive inputs or publish canonical close reports.
- Use an intentional two-page report:
  - Page 1: prominent location, fiscal period, overall disposition, Darden result, control matrix, and open actions.
  - Page 2: repeated location/period header, weekly variances, coverage and status columns, unified exceptions, evidence notes, and page numbering.
- Distinguish `Darden: MATCHED` from the overall close status. Use green only for fully clean closes and amber for `CLOSED WITH REVIEW`.
- Show both largest weekly variance and period-net variance, use accounting currency formatting, and replace internal source-folder names with reader-friendly labels.
- Build follow-up items from every control outcome; “No exceptions” appears only when no review or blocking item exists.
- Export PDF from the canonical workbook through Windows Excel automation so the two formats remain identical. Validate that the PDF opens, contains the location heading, and has the expected page structure; export failure blocks publication and archiving.
- Snapshot and hash the Summary, activity files, Darden PDF, `DLYSYSTT.TXT`, and `TENDER_DETAIL.TXT`. Record canonical archive-relative paths and a close manifest.
- Make closeout transactional: render and verify temporary artifacts, copy and hash-verify all evidence, atomically publish the workbook/PDF, then remove live inputs and prune empty period folders. Partial failures leave original evidence intact.

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
- Finish with two local commits—control/report hardening, then Dropbox/runtime organization—and publish through a GitHub branch and draft pull request.

## Test and Acceptance Plan

- Add negative tests for wrong-store Summary/activity/Darden, duplicate and overlapping weeks, malformed money, period mismatches, missing tender evidence, wrong Micros source, partial coverage, output locking, PDF failure, and archive failure.
- Test scheduled Monday handling for both locations, including a Monday with real activity/POS data.
- Test all status outcomes, including Richmond’s amber result and Virginia Beach’s green result.
- Verify review reports cannot be mistaken for canonical close artifacts.
- Verify source hashes, archive paths, cleanup idempotency, shared-inbox routing, multi-store runs, and nonzero wrapper exit codes.
- Keep Linux CI green with PDF automation mocked; run Windows Excel/PDF integration tests locally.
- Reconcile the real June sources independently, regenerate both artifacts, visually inspect both pages, and verify Dropbox hashes after consolidation.

## Assumptions

- Both locations use Monday as their recurring scheduled closed weekday, subject to evidence-based open-day detection.
- The amber limit is `$5.00` absolute for each weekly control and its period aggregate; any larger value blocks close.
- March 2026 material remains archived historical evidence but is not used in current logic or report narrative.
- Workbook and PDF are both required monthly deliverables.
