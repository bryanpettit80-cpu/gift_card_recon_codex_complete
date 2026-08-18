# Gift Card Workspace Relocation Runbook

This runbook moves the complete Gift Card Reconciliation operator workspace from
the Dropbox root to the master automation folder without moving its external
Micros exports.

## Target layout

```text
Dropbox/
  Automations/
    Gift Card Reconciliation/
      Run Weekly Gift Card Reconciliation.cmd
      Run Monthly Gift Card Close.cmd
      01 Weekly Gift Card Activity Reports/
      02 Monthly Close Inputs/
      03 Finished Reports/
      04 Archive/
      _automation_runs/
      Gift Card Reconciliation Automation/
  micros_data/RC-Richmond-current/
  GETLinkedData-VB/
```

The workspace is one move unit. Do not move only the inner program folder, and
do not move either external Micros folder.

## Preconditions

1. Deploy a version containing explicit Dropbox-root support and
   relocation-safe weekly duplicate validation.
2. Run `_program\run_tests.ps1` from that program version.
3. Confirm both weekly Activity inboxes and the shared Darden inbox are not being
   changed, no reconciliation process is running, and Dropbox is fully synced.
4. Confirm `Dropbox\Automations\Gift Card Reconciliation` does not exist.
5. Record the source file count, total bytes, and SHA-256 inventory outside the
   folder being moved.

## Move

Move this exact source as one directory:

```text
Dropbox\Gift Card Reconciliation
```

to this exact destination:

```text
Dropbox\Automations\Gift Card Reconciliation
```

Do not leave a junction, shortcut, duplicate program tree, or second writable
operations root at the old location.

## Verification

1. Recompute the destination inventory and require every relative path, size,
   and SHA-256 hash to match the source inventory.
2. Confirm the parent launchers still resolve the nested program and pass their
   own folder as `-OperationsRoot`.
3. Confirm the runners resolve `-DropboxRoot` to the Dropbox directory above
   `Automations`.
4. Confirm Richmond resolves to `micros_data\RC-Richmond-current` and Virginia
   Beach resolves to `GETLinkedData-VB`, both at the Dropbox root.
5. Run setup without `-SkipInstall`. The program-root fingerprint must rebuild
   the external editable runtime for the new source path.
6. Run the isolated test suite. Exercise archived-week duplicate and monthly
   archive-reissue checks against copies or test fixtures, not live evidence.
7. Wait until Dropbox reports the move fully synchronized before normal use.

Weekly schema-1 manifests intentionally retain their original absolute paths as
historical provenance. The program does not rewrite them after relocation. It
selects the current canonical workbook from the trusted operations layout and
validates that workbook against the contained archive record's relative path,
name, size, and SHA-256 hash.

## Rollback

Before any new live run, a failed path or hash verification can be rolled back
by moving the complete workspace to its original Dropbox-root location. After a
new live run, stop and reconcile newly produced files before attempting any
rollback; do not merge two writable workspace copies.
