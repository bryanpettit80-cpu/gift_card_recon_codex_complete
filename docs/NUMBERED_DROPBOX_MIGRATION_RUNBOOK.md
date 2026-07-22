# Numbered Dropbox Layout Migration Runbook

The numbered-layout migration is implemented by
`_program\maintenance\migrate_to_numbered_layout.ps1`. It is read-only unless
`-Apply` or `-Rollback` is supplied.

## What the tool moves

Only the following legacy business-data roots are inventoried:

- `9354 - Weekly` and `9355 - Weekly`
- `Monthly Close`
- `Output`
- `Archive - Old Files`
- legacy `input` and `reports` folders, when present

Program files, `.git`, launchers, `_program`, the nested automation repository,
caches, build output, and temporary files are excluded. Historic archive files
move from `Archive - Old Files\<internal path>` to
`04 Archive\<same internal path>` without changing names or bytes.
Migration-generated `*.pre.json`, `*.post.json`, and `*.rollback.json` files in
`04 Archive\Cleanup Manifests` are operational evidence, not business inputs;
they are excluded from every later inventory and plan fingerprint.

## Dry run and approval fingerprint

Before the reviewed dry run, stop both Gift Card launchers, close Excel, and
confirm Dropbox is not adding an Activity or Darden file to the legacy folders.
Keep the workspace quiescent through Apply, Verify, nested-checkout installation,
and operator-asset deployment. If any legacy business file appears after the
inventory, postflight blocks completion and requires a new reviewed fingerprint.

Run the dry run from the current clean merged checkout and explicitly identify
the parent operations folder:

```powershell
$root = "C:\Users\bryan\Dropbox\Gift Card Reconciliation"
$json = & .\_program\maintenance\migrate_to_numbered_layout.ps1 -OperationsRoot $root
$plan = $json | ConvertFrom-Json
$plan.summary
$plan.plan_sha256
```

Dry run does not create folders or manifests. Review every entry in
`$plan.files`, especially `source_relative`, `destination_relative`, hashes,
and the zero-conflict summary. The plan fingerprint changes if any inventoried
file, destination, or mapping changes.

## Apply and verify

Apply with the reviewed fingerprint:

```powershell
& .\_program\maintenance\migrate_to_numbered_layout.ps1 `
  -OperationsRoot $root `
  -Apply `
  -ExpectedPlanSha256 $plan.plan_sha256
```

Apply performs a full conflict and lock preflight before its first change. It
writes timestamped `*.pre.json` and checkpointed `*.post.json` manifests under
`04 Archive\Cleanup Manifests`. Each destination is copied to a partial file,
SHA-256 verified, atomically published, and verified again before the source is
quarantined and removed. A failure leaves either the original or two verified
copies and records `blocked` plus the exact error in the post-migration manifest.

If Apply stops after its immutable preflight write, or after a file operation
reaches a checkpointed source/quarantine/destination state, run the same command
again with the same `-ExpectedPlanSha256`. The tool finds the matching
unfinished checkpoint, verifies its integrity, and validates every current
source, quarantine, and destination before making another change. It then
continues that original reviewed plan and checkpoint instead of building a new
plan from the partially migrated layout. A changed or unplanned business file,
a tampered checkpoint, or ambiguous matching checkpoints blocks the resume
before mutation. A hard stop in the middle of copying can leave a
`.gc-layout-*.partial` file; the tool intentionally blocks on that orphan for
manual review instead of guessing that an incomplete copy is safe to remove.

Verify the completed post-migration manifest:

```powershell
$post = Get-ChildItem "$root\04 Archive\Cleanup Manifests\*.post.json" |
  Sort-Object LastWriteTimeUtc |
  Select-Object -Last 1

& .\_program\maintenance\migrate_to_numbered_layout.ps1 `
  -OperationsRoot $root `
  -Verify `
  -ManifestPath $post.FullName
```

Re-running `-Apply` is idempotent. Files already in the numbered layout are
hash-verified, matching old/new duplicates are reduced to the verified
destination, and different destination content blocks the complete preflight.

## Install the nested program-only checkout

Perform this cutover only after the implementation PR is merged and the live
post-migration manifest verifies successfully:

1. Save both a `git bundle --all` and a `git archive` of the clean outer checkout
   beneath `%LOCALAPPDATA%\GiftCardRecon\layout-migration-backup\<timestamp>`.
2. Make the nested-checkout destination empty. Do not run `git clone` over an
   existing deployment snapshot. If `Gift Card Reconciliation Automation`
   exists without `.git`, verify every file listed in its deployment manifest,
   then move the complete snapshot to a unique sibling backup:

```powershell
$program = Join-Path $root "Gift Card Reconciliation Automation"
if (Test-Path -LiteralPath $program) {
  if (Test-Path -LiteralPath (Join-Path $program ".git")) {
    throw "A Git checkout already occupies $program; validate and reuse it instead of cloning over it."
  }

  $manifestPath = Join-Path $program "deployment-manifest.json"
  if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
    throw "The existing deployment has no manifest and must be reviewed manually: $program"
  }
  $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
  $manifestFiles = @($manifest.files)
  if ([string]$manifest.project -ne "gift_card_recon_codex_complete" -or
      $manifestFiles.Count -eq 0 -or
      [long]$manifest.file_count -ne $manifestFiles.Count) {
    throw "The existing deployment manifest identity or file count is invalid."
  }
  $programFull = [IO.Path]::GetFullPath($program)
  $programPrefix = $programFull.TrimEnd([char[]]"\/") + [IO.Path]::DirectorySeparatorChar
  foreach ($file in $manifestFiles) {
    $relative = [string]$file.path
    if ([string]::IsNullOrWhiteSpace($relative) -or
        [IO.Path]::IsPathRooted($relative) -or
        @($relative -split "[\\/]" | Where-Object { $_ -eq ".." }).Count -gt 0) {
      throw "Unsafe deployment-manifest path: $relative"
    }
    $candidate = [IO.Path]::GetFullPath((Join-Path $programFull $relative))
    if (-not $candidate.StartsWith($programPrefix, [StringComparison]::OrdinalIgnoreCase) -or
        -not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
      throw "Missing or out-of-root deployed file: $relative"
    }
    $actual = Get-FileHash -LiteralPath $candidate -Algorithm SHA256
    if ($actual.Hash -ne ([string]$file.sha256) -or
        (Get-Item -LiteralPath $candidate).Length -ne [long]$file.bytes) {
      throw "Existing deployment does not match its manifest: $relative"
    }
  }

  $stamp = [DateTime]::UtcNow.ToString("yyyyMMddTHHmmssZ")
  $snapshotBackup = Join-Path $root ".Gift Card Reconciliation Automation.pre-nested-$stamp"
  if (Test-Path -LiteralPath $snapshotBackup) {
    throw "Snapshot backup path is occupied: $snapshotBackup"
  }
  Move-Item -LiteralPath $program -Destination $snapshotBackup
  if ((Test-Path -LiteralPath $program) -or
      -not (Test-Path -LiteralPath (Join-Path $snapshotBackup "deployment-manifest.json") -PathType Leaf)) {
    throw "The prior deployment snapshot was not moved safely."
  }
}
```

   Keep `$snapshotBackup` until the nested checkout, tests, and operator assets
   are verified. Do not delete an unverified or unmanifested destination.
3. Clone the private repository's `main` branch into the now-absent destination,
   verify its `HEAD` equals `origin/main`, and initialize its runtime while
   running the tests. A new checkout must not use `-SkipInstall`:

```powershell
git clone --branch main --single-branch `
  https://github.com/bryanpettit80-cpu/gift_card_recon_codex_complete.git `
  $program
if ($LASTEXITCODE -ne 0) { throw "Nested checkout clone failed with exit code $LASTEXITCODE." }
git -C $program fetch origin main
if ($LASTEXITCODE -ne 0) { throw "Nested checkout fetch failed with exit code $LASTEXITCODE." }
if ((git -C $program rev-parse HEAD) -ne (git -C $program rev-parse origin/main)) {
  throw "Nested checkout does not match origin/main."
}
& "$program\_program\run_tests.ps1"
if ($LASTEXITCODE -ne 0) { throw "Nested checkout tests failed with exit code $LASTEXITCODE." }
```

4. Only after that clone passes, remove the old outer `.git` and outer tracked
   program files. Do not remove numbered business folders or `_automation_runs`.
5. Deploy the parent launchers, START HERE guide, drop notes, and required
   operator folders from the nested checkout's tracked templates:

```powershell
& "$root\Gift Card Reconciliation Automation\_program\install_operator_assets.ps1" `
  -OperationsRoot $root
```

The installer stages and SHA-256-verifies the complete managed asset set before
changing live files, backs up the prior set, publishes the new set as one
transaction, and restores every prior file if a late replacement fails. It also
retires the exact obsolete health launcher released by the prior deployment;
an unrecognized same-name file is preserved with a warning.

Finally, update the Codex/GitHub audit discovery root and Codex trusted-project
entry to the nested checkout, then confirm a targeted search finds no active
configuration still pointing at the former outer repository.

## Rollback

Rollback uses the post-migration manifest as its authority:

```powershell
& .\_program\maintenance\migrate_to_numbered_layout.ps1 `
  -OperationsRoot $root `
  -Rollback `
  -ManifestPath $post.FullName
```

Rollback first verifies every applicable destination. It restores each old
source path and hash. A destination created by that migration is removed only
after the restored source verifies; a destination that predated migration is
preserved. A separate `*.rollback.json` report records the result.

## Validation

Run the isolated fixture suite without touching Dropbox business data:

```powershell
& .\_program\maintenance\test_migrate_to_numbered_layout.ps1
& .\_program\maintenance\test_install_operator_assets.ps1
```

The fixture covers dry run, Apply, resume after checkpoint creation, resume
after a completed file move, manifest verification, generated-manifest
exclusion, idempotent reapply, rollback, archive-path preservation, conflicting
destinations, locked sources, and program-file exclusion. The older
`consolidate_dropbox.ps1` is retained only as the historical July 11
consolidator and intentionally refuses to run after `04 Archive` exists.
The operator-assets fixture verifies a complete successful refresh, safe stale
launcher retirement, preservation of unrelated files, and restoration of the
entire prior managed set after a late locked-file failure.
