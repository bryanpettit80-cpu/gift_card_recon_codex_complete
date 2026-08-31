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

The Dropbox setup folder is an **untrusted payload source**, not a release authority.
`Install-DailyGiftCardCopyTask.cmd` in that folder is deliberately disabled; do not use
it to install, repair, or neutralize the task. The RESSERVER task account must instead
have a locally provisioned trusted verifier at:

```text
%LOCALAPPDATA%\GiftCardRecon\RichmondMicrosExport\Trusted-InstallDailyGiftCardCopyTask.ps1
```

Provision that local program asset from the reviewed canonical release through a
separate trusted local release process, not by copying from the synced Dropbox setup
folder. The verifier embeds the approved payload filename, byte length, and SHA-256;
do not create or commit a separate release manifest. The trusted local directory, its
`GiftCardRecon` parent, and the verifier must grant write access only to the RESSERVER
task account, `SYSTEM`, and local Administrators. The verifier rejects reparse points
and any other writable principal.

Run the protected local verifier itself, using the synced payload only as a data input:

```powershell
& "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\GiftCardRecon\RichmondMicrosExport\Trusted-InstallDailyGiftCardCopyTask.ps1" -SourceScript "C:\Users\customer\Dropbox\micros_data\RC-Richmond\_gift_card_current_export_setup\Copy-GiftCardExportToDropbox.cmd"
```

Before reading the Dropbox payload, the protected local verifier replaces the
same-named task with a harmless local no-op. It then checks the payload filename, byte
length, and SHA-256 against its embedded reviewed release metadata **before** staging.
It rechecks the staged file and the final local task payload before registering the
daily action:

```text
%LOCALAPPDATA%\GiftCardRecon\RichmondMicrosExport\payload\Copy-GiftCardExportToDropbox.cmd
```

The installed task never executes the mutable Dropbox setup copy. If the protected
verifier is missing, has unsafe ACLs, is a reparse point, rejects the payload, or task
activation fails, the harmless action remains in place. If the initial
neutralization itself fails, the verifier reports a `SECURITY ERROR`; correct or disable
the existing task manually before retrying.

For an approved payload change, update the verifier's embedded release metadata through
review, provision that reviewed verifier through the trusted local process, and then
rerun it. A Dropbox sync alone must never refresh the local trust anchor. The repository
copy of `Trusted-InstallDailyGiftCardCopyTask.ps1` is the release input for that
controlled provisioning step; the trusted local verifier does not trust any Dropbox
copy of itself.

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

The launcher resolves the user's Dropbox root and maps store `9354` to
`micros_data\RC-Richmond-current` there. After the operations workspace moves
below `Dropbox\Automations`, the Micros export remains at the Dropbox root. Use
the launcher's `-DropboxRoot` option if the root cannot be inferred.
