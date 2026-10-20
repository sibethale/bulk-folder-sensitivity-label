# Bulk Folder Sensitivity Label Changer

`Set-FolderSensitivityLabel.ps1` applies, changes, or removes a Microsoft Purview sensitivity
label on every file in a folder (optionally recursive), using the Microsoft Purview Information
Protection client module.

## Files

| File | Purpose |
|---|---|
| `Set-FolderSensitivityLabel.ps1` | The main script. |
| `Run-LabelChanger.cmd` | Launcher that starts the script under Windows PowerShell 5.1 and forwards all arguments. |
| `Get-SensitivityLabelList.ps1` | Lists your tenant's labels with their GUIDs (via Security & Compliance PowerShell). |
| `test-run.ps1` | Self-test using stubbed labelling cmdlets. Touches no real files. |

## IMPORTANT: use Windows PowerShell 5.1, not PowerShell 7

On this machine the client module (`PurviewInformationProtection` 3.2.92.0, installed with the
Purview Information Protection client at `C:\Program Files (x86)\Microsoft Purview Information
Protection`) loads **only** under `powershell.exe` (5.1). Under `pwsh` 7 the import fails with:

```
Could not load file or assembly 'MSIP.Scanner, Version=3.2.0.0, ...'
```

Verified: 5.1 (both 32- and 64-bit) imports all 49 cmdlets; pwsh 7 fails. The script therefore
blocks itself with a clear message if started in PowerShell 7. Use either:

```powershell
powershell.exe -NoProfile -File .\Set-FolderSensitivityLabel.ps1 -Path 'C:\Data' -Recurse -ShowCurrentLabels
```
or simply:
```
Run-LabelChanger.cmd -Path "C:\Data" -Recurse -ShowCurrentLabels
```

Note: when passing `-Confirm:$false` through `powershell.exe`, use `-Command` rather than `-File`
(`-File` turns every argument into a plain string).

## Prerequisites

1. Windows PowerShell 5.1.
2. The Microsoft Purview Information Protection client installed (it ships the module). Confirmed
   installed here.
3. Sign-in. Interactive runs use your Office identity. For unattended/scheduled runs:
   ```powershell
   Set-Authentication -AppId <app-id> -AppSecret <secret> -TenantId <tenant-id> -DelegatedUser <upn>
   ```
4. Rights to apply the label, plus write access to the files.

## Typical workflow

### 1. Find your label GUIDs

```powershell
.\Get-SensitivityLabelList.ps1
```
Connects to Security & Compliance PowerShell and prints every label with its GUID.

You can also read labels already on disk — this is read-only and changes nothing:
```
Run-LabelChanger.cmd -Path "C:\Data\Reports" -Recurse -ShowCurrentLabels
```
It prints a per-label count and the label GUIDs found.

### 2. Dry run

```
Run-LabelChanger.cmd -Path "C:\Data\Reports" -Recurse -LabelId "bd0b5b6a-...-1f1a2b3c4d5e" -WhatIf
```

### 3. Do it

```
Run-LabelChanger.cmd -Path "C:\Data\Reports" -Recurse -LabelId "bd0b5b6a-...-1f1a2b3c4d5e" -Include "*.docx","*.xlsx","*.pptx","*.pdf" -Exclude "~$*" -PreserveFileDetails
```
You get one confirmation prompt for the whole batch (`-Confirm:$false` to suppress), a progress
bar, and a CSV report next to the script.

### Removing or downgrading a label

Policy usually requires a justification when lowering or removing a label:

```
Run-LabelChanger.cmd -Path "C:\Data\Old" -Recurse -RemoveLabel -JustificationMessage "Content reviewed and declassified"
```

Downgrade (e.g. Confidential -> General) = apply the lower label plus a justification:

```
Run-LabelChanger.cmd -Path "C:\Data\Old" -Recurse -LabelId "<lower-label-guid>" -JustificationMessage "Reviewed, no longer confidential"
```

## Parameters

| Parameter | Purpose |
|---|---|
| `-Path` | Folder to process (required). |
| `-LabelId` | GUID of the label to apply. |
| `-RemoveLabel` | Remove the current label instead of applying one. |
| `-ShowCurrentLabels` | Report-only; changes nothing. |
| `-Recurse` | Include sub-folders. |
| `-Include` / `-Exclude` | Wildcard filename filters. |
| `-JustificationMessage` | Justification for downgrade/removal. |
| `-Owner` | Rights Management owner to stamp (service-principal runs). |
| `-PreserveFileDetails` | Don't change Last Modified / Modified By. |
| `-Force` | Re-apply even if the file already has the target label. |
| `-ReportPath` | Where to write the CSV report. |
| `-RetryCount` | Retries per file on transient errors (default 2). |
| `-ThrottleDelayMs` | Delay between files, to go easy on the service. |
| `-ShowAll` | Include unchanged files in the on-screen table. |
| `-MaxDisplayRows` | Cap on on-screen rows (default 50; 0 = unlimited). |
| `-NoVerify` | Skip the post-change read-back of each file's label. |
| `-PassThru` | Also emit result objects to the pipeline. |
| `-WhatIf` / `-Confirm` | Standard PowerShell safety switches. |
| `-SkipModuleCheck` | Testing only; bypasses module loading. |

## Report

### On screen

Every run prints a details table of what changed, with the label before and after:

```
Details

File       Was    Now     Status  Message
----       ---    ---     ------  -------
demo1.docx (none) General Success Label 00000000-0000-0000-0000-000000000000 applied
demo2.docx (none) General Success Label 00000000-0000-0000-0000-000000000000 applied

Summary
  Success    2
```

The `Now` column is read back from the file after the change, so it reflects what the file really
reports — not just what was requested. Use `-NoVerify` to skip that extra read.

By default unchanged files are hidden and the table is capped at 50 rows:

| Switch | Effect |
|---|---|
| `-ShowAll` | Also list skipped/unchanged files |
| `-MaxDisplayRows 0` | No row cap |
| `-MaxDisplayRows 10` | Cap at 10 rows |
| `-NoVerify` | Don't re-read files after changing them |
| `-PassThru` | Also emit result objects to the pipeline |

### CSV

Every run also writes `LabelReport-<timestamp>.csv` with `File, PreviousLabel, PreviousLabelId,
NewLabel, NewLabelId, Action, Status, Message` — always every file, regardless of what's on screen.

| Status | Meaning |
|---|---|
| `Success` | Label applied/removed and verified |
| `Skipped` | Already had the target label, or had none to remove |
| `Reported` | Read-only mode |
| `WhatIf` | Preview only |
| `Error` | The labelling call failed — see `Message` |
| `Mismatch` | The call reported success but the file doesn't carry the expected label |

Service errors are captured per file rather than aborting the run, e.g.:

```
Status : Error
Message: Can't find label ID: 00000000-.... Verify the label is enabled
```

### About `Mismatch`

Confirmed by testing: `.txt` and `.csv` files accept the `Set-FileLabel` call without error, but
store nothing — plain-text formats have nowhere to keep the metadata unless the label also applies
encryption (which produces a `.ptxt`/`.pfile`). The script catches this and reports `Mismatch`
instead of falsely claiming success. Office formats (`.docx`, `.xlsx`, `.pptx`) and `.pdf` store
labels natively and report `Success`.

The same applies to `.ps1`, `.md`, `.json`, `.log` and any other plain-text file. If you need
those protected, pick a label that applies encryption — otherwise leave them out of `-Include`.

### "It didn't label one of my files"

The most common cause is `-Include`/`-Exclude`. The script now prints a line such as:

```
Found 3 file(s) under '...'.
Excluded by -Include/-Exclude: 1 file(s) -> .ps1 x1
```

so you can see at a glance that the folder held more files than were processed. Run with
`-ShowCurrentLabels` and no `-Include` to see everything the folder contains.

### "Is it stuck, or just slow?"

Large runs over OneDrive/SharePoint-synced folders are slow because every cloud-only file must be
downloaded in full before its label can be read. Seconds per file is normal.

The progress bar shows a counter, elapsed time and a projected finish so you can tell the
difference without guessing:

```
Processing sensitivity labels
File 118 of 244 | elapsed 00:04:12 | ~00:04:29 left (2.1s/file)
C:\Data\Reports\03 - Design.docx
```

To confirm from a second window, without disturbing the run:

```powershell
# CPU must increase between two samples ~20s apart
Get-Process powershell | Format-Table Id, CPU, StartTime

# The MIP SDK's own log should have been written to within the last minute
(Get-Item "$env:LOCALAPPDATA\Microsoft\MSIP\mip\powershell.exe\mip\logs\mip_sdk.miplog").LastWriteTime
```

Frozen CPU *and* a stale log means it has genuinely hung. Otherwise it's working.

The end-of-run `Summary` block also reports total `Duration` and the average seconds per file.

## Validation performed

* `test-run.ps1` passes all 7 scenarios (report, WhatIf, apply, skip-already-labelled, force,
  remove, invalid input) under both PowerShell 5.1 and 7 with stubbed cmdlets.
* Live read-only run against the real module over a temp folder: succeeded.
* Live apply with a deliberately invalid GUID: the service error was captured per file and written
  to the CSV, no crash.
* Live apply of a real tenant label to throwaway `.docx` files in
  `%TEMP%`: applied and verified, `Now` column resolved to `General`. Re-running the same command
  correctly skipped both files as `Already has target label`.
* Live apply to `.txt`/`.csv`: correctly reported `Mismatch` rather than a false success.

## Notes and gotchas

* Only file types supported by Purview labelling can be labelled natively; others are protected as
  `.pfile` or fail — check the CSV.
* Files open in Office (lock files like `~$report.docx`) will fail; exclude them with `-Exclude "~$*"`.
* Files already carrying the target label are skipped unless `-Force` is used.
* Labelling protected/encrypted content requires rights to that content.
* For content already in SharePoint/OneDrive, prefer a Purview auto-labelling policy over running
  this against a synced folder.

### Harmless messages you can ignore

* `Error removing log file: ...\mip_sdk_backup.miplog` — the MIP SDK failing to rotate its own
  log. It has nothing to do with your files and does not affect results.
* A warning about `.849C9593-D756-4E56-8D6E-42412F2A707B` being *used by another process* — that
  is OneDrive's internal sync marker, locked by OneDrive. It is reported and skipped.
