# Richmond Micros Gift Card Export

Richmond store `9354` uses current Micros exports from `RESSERVER`.

The active server source folder is:

```text
C:\GetLinkedData
```

The gift-card monthly close expects the current files to sync back to Bryan's PC at:

```text
C:\Users\bryan\Dropbox\micros_data\RC-Richmond-current
```

The server-side scripts in this folder publish only the files needed by monthly close:

```text
Micros3700.7z
DLYSYSTT.TXT
TENDER_DETAIL.TXT
```

On `RESSERVER`, the installed task is:

```text
Gift Card Export Copy to Dropbox
```

Run `Install-DailyGiftCardCopyTask.cmd` from the synced setup folder. Before reading any
synced file, the installer replaces the same-named task with a harmless local no-op.
This fail-closed step prevents an older task from continuing to execute a mutable
Dropbox copy if staging fails. The installer then copies
`Copy-GiftCardExportToDropbox.cmd` to the server user's private local application-data
folder, verifies the installed copy by SHA-256, and configures the task to execute that
local snapshot:

```text
%LOCALAPPDATA%\GiftCardRecon\RichmondMicrosExport\Copy-GiftCardExportToDropbox.cmd
```

The installed task never executes the mutable Dropbox setup copy. If staging or
verification fails, the harmless action remains in place. If the initial neutralization
itself fails, the installer reports a `SECURITY ERROR`; correct or disable the existing
task manually before retrying. Rerun the installer after an approved script update to
refresh and re-verify the private snapshot.

It runs daily at `06:35`, after the normal GetLinked export, and copies files into:

```text
C:\Users\customer\Dropbox\micros_data\RC-Richmond-current
```

To confirm installation on the server, run:

```text
Check-GiftCardExportSetup.cmd
```

It writes:

```text
C:\Users\customer\Dropbox\micros_data\RC-Richmond-current\GiftCardSetupStatus.txt
```

Use the folder path for monthly close:

```powershell
& ".\Run Monthly Gift Card Close.cmd" -Store 9354 -Period FY27-M01
```

The launcher automatically maps store `9354` to `..\micros_data\RC-Richmond-current`.
