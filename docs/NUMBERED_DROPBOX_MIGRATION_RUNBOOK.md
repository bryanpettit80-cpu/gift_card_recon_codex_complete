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

Program files, `.git`, launchers, `_program`, the deployed automation snapshot,
caches, build output, and temporary files are excluded. Historic archive files
move from `Archive - Old Files\<internal path>` to
`04 Archive\<same internal path>` without changing names or bytes.
Migration-generated `*.pre.json`, `*.post.json`, and `*.rollback.json` files in
`04 Archive\Cleanup Manifests` are operational evidence, not business inputs;
they are excluded from every later inventory and plan fingerprint.

## Dry run and approval fingerprint

Before the reviewed dry run, stop both Gift Card launchers, close Excel, and
confirm Dropbox is not adding an Activity or Darden file to the legacy folders.
Keep the workspace quiescent through Apply, Verify, program deployment,
and operator-asset deployment. If any legacy business file appears after the
inventory, postflight blocks completion and requires a new reviewed fingerprint.

Run the dry run from the current clean merged checkout and explicitly identify
the parent operations folder:

```powershell
$root = Join-Path $env:USERPROFILE "Dropbox\Gift Card Reconciliation"
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
copies and records `blocked` plus the exact error in the post manifest.

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

Verify the completed post manifest:

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

## Deploy the program snapshot

Perform this cutover only after the implementation PR is merged and the live
migration post manifest verifies successfully. GitHub remains the authoritative
source, the local clone remains the sole Git working copy, and Dropbox receives
only a verified deployment with no `.git` directory:

1. In the local clone, verify `HEAD` equals `origin/main`, the tracked worktree
   is clean, and `_program\run_tests.ps1` passes.
2. Deploy the program and operator assets from that local clone:

```powershell
$repo = Join-Path $env:USERPROFILE "Documents\Repos\gift_card_recon_codex_complete"
& "$repo\_program\maintenance\deploy_operator_program.ps1" -OperationsRoot $root
```

3. Verify `Gift Card Reconciliation Automation\deployment-manifest.json` exists,
   the deployed revision matches the intended local commit, and
   `Gift Card Reconciliation Automation\.git` does not exist.
4. Run `Check Gift Card Reconciliation Health.cmd` from the operations root and
   resolve every blocking result before returning the launchers to operators.

Finally, update Codex/GitHub discovery and trusted-project configuration to the
local clone, then confirm no active configuration treats the Dropbox deployment
as a Git checkout.

## Rollback

Rollback uses the migration post manifest as its authority:

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
```

The fixture covers dry run, Apply, resume after checkpoint creation, resume
after a completed file move, manifest verification, generated-manifest
exclusion, idempotent reapply, rollback, archive-path preservation, conflicting
destinations, locked sources, and program-file exclusion. The older
`consolidate_dropbox.ps1` is retained only as the historical July 11
consolidator and intentionally refuses to run after `04 Archive` exists.
