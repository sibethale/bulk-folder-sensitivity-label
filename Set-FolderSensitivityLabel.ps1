<#
.SYNOPSIS
    Bulk-applies (or removes) a Microsoft Purview sensitivity label on every file in a folder.

.DESCRIPTION
    Wraps the PurviewInformationProtection PowerShell module (Get-FileStatus / Set-FileLabel /
    Remove-FileLabel) so an entire folder tree can be relabelled in one pass, with filtering,
    skip-if-already-correct logic, retries, a CSV report and full -WhatIf / -Confirm support.

    Requires:
      * Windows PowerShell 5.1 (powershell.exe). The Purview client's AIP.dll cannot resolve its
        dependencies under PowerShell 7 (pwsh), so run this script with powershell.exe.
      * The Microsoft Purview Information Protection client installed, which provides the
        PurviewInformationProtection module.
      * A signed-in labelling identity. Interactive use picks up your Office sign-in; for
        unattended runs call Set-Authentication once with a service principal.

.PARAMETER Path
    Folder whose files should be relabelled.

.PARAMETER LabelId
    GUID of the sensitivity label to apply. Discover GUIDs with -ShowCurrentLabels on files that
    already carry the label, or with Get-Label in Security & Compliance PowerShell.

.PARAMETER RemoveLabel
    Remove the existing label instead of applying one.

.PARAMETER ShowCurrentLabels
    Report-only mode: lists the current label of every matching file and changes nothing.

.PARAMETER Recurse
    Include sub-folders.

.PARAMETER Include
    Wildcard file filters, e.g. '*.docx','*.xlsx','*.pdf'. Default: all files.

.PARAMETER Exclude
    Wildcard filters for files to skip, e.g. '~$*','*.tmp'.

.PARAMETER JustificationMessage
    Required by policy when lowering or removing a label.

.PARAMETER Owner
    Sets the Rights Management owner of the labelled file (useful for service-principal runs).

.PARAMETER PreserveFileDetails
    Keep the file's Last Modified / Modified By metadata unchanged.

.PARAMETER Force
    Relabel files that already carry the target label (skipped by default).

.PARAMETER ReportPath
    CSV file for the per-file result. Defaults to a timestamped file next to the script.

.PARAMETER RetryCount
    Extra attempts per file on transient failures (default 2).

.PARAMETER ThrottleDelayMs
    Pause between files, in milliseconds (default 0).

.PARAMETER MaxDisplayRows
    Maximum rows in the on-screen details table (default 50). Use 0 for no limit.

.PARAMETER ShowAll
    Include unchanged/skipped files in the on-screen table. By default only files that were
    changed, previewed or failed are listed; the CSV always contains every file.

.PARAMETER NoVerify
    Skip the post-change re-read of each file's label. Faster, but the "Now" column then shows the
    intended label rather than the confirmed one.

.PARAMETER PassThru
    Also emit the result objects to the pipeline, for further filtering or custom formatting.

.PARAMETER SkipModuleCheck
    Skip loading PurviewInformationProtection. Only for testing with stubbed cmdlets.

.EXAMPLE
    # Always start powershell.exe (5.1), not pwsh:
    .\Set-FolderSensitivityLabel.ps1 -Path 'C:\Data\Reports' -Recurse -ShowCurrentLabels

.EXAMPLE
    .\Set-FolderSensitivityLabel.ps1 -Path 'C:\Data\Reports' -Recurse -LabelId 'bd0b5b6a-9d3b-4f2a-8b1d-1f1a2b3c4d5e' -WhatIf

.EXAMPLE
    .\Set-FolderSensitivityLabel.ps1 -Path 'C:\Data\Reports' -Recurse -LabelId 'bd0b5b6a-9d3b-4f2a-8b1d-1f1a2b3c4d5e' -Include '*.docx','*.xlsx','*.pdf' -Exclude '~$*' -PreserveFileDetails

.EXAMPLE
    .\Set-FolderSensitivityLabel.ps1 -Path 'C:\Data\Old' -Recurse -RemoveLabel -JustificationMessage 'Content reviewed and declassified'
#>
[CmdletBinding(DefaultParameterSetName = 'Apply', SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$Path,

    [Parameter(Mandatory = $true, ParameterSetName = 'Apply')]
    [ValidatePattern('^\{?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}\}?$')]
    [string]$LabelId,

    [Parameter(Mandatory = $true, ParameterSetName = 'Remove')]
    [switch]$RemoveLabel,

    [Parameter(Mandatory = $true, ParameterSetName = 'Report')]
    [switch]$ShowCurrentLabels,

    [switch]$Recurse,

    [string[]]$Include,

    [string[]]$Exclude,

    [string]$JustificationMessage,

    [string]$Owner,

    [switch]$PreserveFileDetails,

    [Parameter(ParameterSetName = 'Apply')]
    [switch]$Force,

    [string]$ReportPath,

    [ValidateRange(0, 10)]
    [int]$RetryCount = 2,

    [ValidateRange(0, 60000)]
    [int]$ThrottleDelayMs = 0,

    [ValidateRange(0, 100000)]
    [int]$MaxDisplayRows = 50,

    [switch]$ShowAll,

    [switch]$NoVerify,

    [switch]$PassThru,

    [switch]$SkipModuleCheck
)

$ErrorActionPreference = 'Stop'

function Initialize-LabelingModule {
    if ($SkipModuleCheck) {
        Write-Verbose 'SkipModuleCheck set - assuming labelling cmdlets are already available.'
        return
    }

    if ($PSVersionTable.PSEdition -eq 'Core') {
        throw 'Run this script in Windows PowerShell 5.1 (powershell.exe). The Purview client module fails to load under PowerShell 7 (pwsh) with "Could not load file or assembly MSIP.Scanner". Example: powershell.exe -NoProfile -File .\Set-FolderSensitivityLabel.ps1 -Path C:\Data -Recurse -ShowCurrentLabels'
    }

    if (-not (Get-Module -ListAvailable -Name PurviewInformationProtection)) {
        throw 'The PurviewInformationProtection module was not found. Install the Microsoft Purview Information Protection client (it ships the module), or run: Install-Module -Name PurviewInformationProtection -Scope CurrentUser'
    }

    try {
        Import-Module PurviewInformationProtection -ErrorAction Stop
    }
    catch {
        throw ('Could not load PurviewInformationProtection: {0}' -f $_.Exception.Message)
    }

    foreach ($cmd in 'Get-FileStatus', 'Set-FileLabel', 'Remove-FileLabel') {
        if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
            throw "Required cmdlet '$cmd' was not found after importing PurviewInformationProtection."
        }
    }
}

function Get-AllFile {
    param([string]$Root)

    $gciParams = @{
        LiteralPath = $Root
        File        = $true
        Force       = $true
        ErrorAction = 'SilentlyContinue'
    }
    if ($Recurse) { $gciParams.Recurse = $true }

    Get-ChildItem @gciParams
}

function Select-TargetFile {
    param([object[]]$Files)

    $Files | Where-Object {
        $name = $_.Name
        $keep = $true
        if ($Include) { $keep = @($Include | Where-Object { $name -like $_ }).Count -gt 0 }
        if ($keep -and $Exclude) { $keep = @($Exclude | Where-Object { $name -like $_ }).Count -eq 0 }
        $keep
    }
}

function Get-CurrentLabelInfo {
    param([string]$FilePath)

    try {
        $status = Get-FileStatus -Path $FilePath -ErrorAction Stop
    }
    catch {
        return [pscustomobject]@{ LabelId = $null; LabelName = $null; Error = $_.Exception.Message }
    }

    $id = $null
    $name = $null
    if ($status) {
        $props = $status.PSObject.Properties.Name
        if (($props -contains 'SubLabelId') -and $status.SubLabelId) {
            $id = [string]$status.SubLabelId
            if ($props -contains 'SubLabelName') { $name = $status.SubLabelName }
        }
        elseif (($props -contains 'MainLabelId') -and $status.MainLabelId) {
            $id = [string]$status.MainLabelId
            if ($props -contains 'MainLabelName') { $name = $status.MainLabelName }
        }
    }

    [pscustomobject]@{ LabelId = $id; LabelName = $name; Error = $null }
}

function Invoke-WithRetry {
    param(
        [scriptblock]$Action,
        [int]$Attempts
    )

    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            return & $Action
        }
        catch {
            if ($i -ge $Attempts) { throw }
            Write-Verbose ('Attempt {0} failed: {1}. Retrying...' -f $i, $_.Exception.Message)
            Start-Sleep -Seconds ([int][Math]::Min(10, [Math]::Pow(2, $i)))
        }
    }
}

if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
    throw "Folder not found: $Path"
}
$resolvedRoot = (Resolve-Path -LiteralPath $Path).ProviderPath

$normalizedLabelId = $null
if ($PSCmdlet.ParameterSetName -eq 'Apply') {
    $normalizedLabelId = $LabelId.Trim('{', '}').ToLowerInvariant()
}

Initialize-LabelingModule

$allFiles = @(Get-AllFile -Root $resolvedRoot)
$files = @(Select-TargetFile -Files $allFiles)
$filteredOut = @($allFiles | Where-Object { $files -notcontains $_ })

if ($files.Count -eq 0) {
    Write-Warning "No files matched under '$resolvedRoot'."
    if ($filteredOut.Count -gt 0) {
        Write-Warning ("{0} file(s) were excluded by -Include/-Exclude. Extensions present: {1}" -f
            $filteredOut.Count, (($filteredOut | ForEach-Object { if ($_.Extension) { $_.Extension } else { '(none)' } } | Sort-Object -Unique) -join ', '))
    }
    return
}

Write-Host ("Found {0} file(s) under '{1}'." -f $files.Count, $resolvedRoot) -ForegroundColor Cyan

# Make filter-excluded files visible: silently processing fewer files than the folder holds is
# the most common source of "it didn't label everything".
if ($filteredOut.Count -gt 0) {
    $byExt = $filteredOut |
        Group-Object { if ($_.Extension) { $_.Extension.ToLowerInvariant() } else { '(no extension)' } } |
        Sort-Object Count -Descending
    Write-Host ("Excluded by -Include/-Exclude: {0} file(s) -> {1}" -f
        $filteredOut.Count, (($byExt | ForEach-Object { '{0} x{1}' -f $_.Name, $_.Count }) -join ', ')) -ForegroundColor Yellow
}

# One confirmation for the whole batch instead of a prompt per file.
if ($PSCmdlet.ParameterSetName -ne 'Report' -and -not $WhatIfPreference -and $ConfirmPreference -ne 'None') {
    $what = if ($PSCmdlet.ParameterSetName -eq 'Remove') { 'remove the sensitivity label from' } else { "apply label $normalizedLabelId to" }
    if (-not $PSCmdlet.ShouldContinue(("About to {0} up to {1} file(s) under '{2}'. Continue?" -f $what, $files.Count, $resolvedRoot), 'Bulk sensitivity label change')) {
        Write-Warning 'Cancelled by user.'
        return
    }
    $ConfirmPreference = 'None'
}

if (-not $ReportPath) {
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $reportRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $ReportPath = Join-Path -Path $reportRoot -ChildPath "LabelReport-$stamp.csv"
}

$results = [System.Collections.Generic.List[object]]::new()
$index = 0

foreach ($file in $files) {
    $index++
    Write-Progress -Activity 'Processing sensitivity labels' -Status $file.FullName -PercentComplete ([int](($index / $files.Count) * 100))

    $before = Get-CurrentLabelInfo -FilePath $file.FullName
    $record = [pscustomobject]@{
        File            = $file.FullName
        PreviousLabel   = $before.LabelName
        PreviousLabelId = $before.LabelId
        NewLabel        = $before.LabelName
        NewLabelId      = $before.LabelId
        Action          = 'None'
        Status          = 'Skipped'
        Message         = $before.Error
    }

    if ($before.Error) {
        $record.Action = 'Read'
        $record.Status = 'Error'
        $results.Add($record)
        Write-Warning ('{0}: {1}' -f $file.FullName, $before.Error)
        continue
    }

    if ($PSCmdlet.ParameterSetName -eq 'Report') {
        $record.Action = 'Read'
        $record.Status = 'Reported'
        $record.Message = if ($before.LabelId) { 'Labelled' } else { 'No label' }
        $results.Add($record)
        continue
    }

    if ($PSCmdlet.ParameterSetName -eq 'Remove') {
        if (-not $before.LabelId) {
            $record.Message = 'No label to remove'
            $results.Add($record)
            continue
        }

        if ($PSCmdlet.ShouldProcess($file.FullName, 'Remove sensitivity label')) {
            try {
                Invoke-WithRetry -Attempts ($RetryCount + 1) -Action {
                    $p = @{ Path = $file.FullName; ErrorAction = 'Stop' }
                    if ($JustificationMessage) { $p.JustificationMessage = $JustificationMessage }
                    if ($PreserveFileDetails) { $p.PreserveFileDetails = $true }
                    Remove-FileLabel @p | Out-Null
                }
                $record.Action = 'RemoveLabel'
                $record.Status = 'Success'
                $record.Message = 'Label removed'
                if (-not $NoVerify) {
                    $after = Get-CurrentLabelInfo -FilePath $file.FullName
                    $record.NewLabel = $after.LabelName
                    $record.NewLabelId = $after.LabelId
                    if ($after.LabelId) {
                        $record.Status = 'Mismatch'
                        $record.Message = 'Label still present after removal'
                    }
                }
                else {
                    $record.NewLabel = $null
                    $record.NewLabelId = $null
                }
            }
            catch {
                $record.Action = 'RemoveLabel'
                $record.Status = 'Error'
                $record.Message = $_.Exception.Message
                Write-Warning ('{0}: {1}' -f $file.FullName, $_.Exception.Message)
            }
        }
        else {
            $record.Action = 'RemoveLabel'
            $record.Status = 'WhatIf'
        }

        $results.Add($record)
    }
    else {
        if (-not $Force -and $before.LabelId -and $before.LabelId.ToLowerInvariant() -eq $normalizedLabelId) {
            $record.Message = 'Already has target label'
            $results.Add($record)
            continue
        }

        if ($PSCmdlet.ShouldProcess($file.FullName, "Apply sensitivity label $normalizedLabelId")) {
            try {
                Invoke-WithRetry -Attempts ($RetryCount + 1) -Action {
                    $p = @{ Path = $file.FullName; LabelId = $normalizedLabelId; ErrorAction = 'Stop' }
                    if ($JustificationMessage) { $p.JustificationMessage = $JustificationMessage }
                    if ($PreserveFileDetails) { $p.PreserveFileDetails = $true }
                    if ($Owner) { $p.Owner = $Owner }
                    Set-FileLabel @p | Out-Null
                }
                $record.Action = 'SetLabel'
                $record.Status = 'Success'
                $record.Message = "Label $normalizedLabelId applied"
                if (-not $NoVerify) {
                    $after = Get-CurrentLabelInfo -FilePath $file.FullName
                    $record.NewLabel = $after.LabelName
                    $record.NewLabelId = $after.LabelId
                    if (-not $after.LabelId) {
                        $record.Status = 'Mismatch'
                        $record.Message = ("The labelling call succeeded but '{0}' reports no label. File types without native label support (e.g. .txt, .csv) can only be labelled when the label also applies encryption." -f $file.Extension)
                    }
                    elseif ($after.LabelId.ToLowerInvariant() -ne $normalizedLabelId) {
                        $record.Status = 'Mismatch'
                        $record.Message = ('Expected {0} but file reports {1}' -f $normalizedLabelId, $after.LabelId)
                    }
                }
                else {
                    $record.NewLabelId = $normalizedLabelId
                    $record.NewLabel = $null
                }
            }
            catch {
                $record.Action = 'SetLabel'
                $record.Status = 'Error'
                $record.Message = $_.Exception.Message
                Write-Warning ('{0}: {1}' -f $file.FullName, $_.Exception.Message)
            }
        }
        else {
            $record.Action = 'SetLabel'
            $record.Status = 'WhatIf'
        }

        $results.Add($record)
    }

    if ($ThrottleDelayMs -gt 0) { Start-Sleep -Milliseconds $ThrottleDelayMs }
}

Write-Progress -Activity 'Processing sensitivity labels' -Completed

# --- on-screen detail -------------------------------------------------------

$display = if ($ShowAll -or $PSCmdlet.ParameterSetName -eq 'Report') {
    $results
}
else {
    # Default: show what actually happened, not the untouched files.
    @($results | Where-Object { $_.Status -ne 'Skipped' })
}
$display = @($display)

Write-Host ''
if ($display.Count -eq 0) {
    Write-Host 'No files were changed.' -ForegroundColor Yellow
}
else {
    Write-Host 'Details' -ForegroundColor Green
    $shown = @(if ($MaxDisplayRows -gt 0 -and $display.Count -gt $MaxDisplayRows) { $display[0..($MaxDisplayRows - 1)] } else { $display })

    $table = $shown | Select-Object `
        @{ n = 'File'; e = { $_.File.Substring([Math]::Min($_.File.Length, $resolvedRoot.Length)).TrimStart('\') } },
        @{ n = 'Was'; e = { if ([string]::IsNullOrWhiteSpace($_.PreviousLabel)) { if ($_.PreviousLabelId) { $_.PreviousLabelId } else { '(none)' } } else { $_.PreviousLabel } } },
        @{ n = 'Now'; e = { if ([string]::IsNullOrWhiteSpace($_.NewLabel)) { if ($_.NewLabelId) { $_.NewLabelId } else { '(none)' } } else { $_.NewLabel } } },
        Status,
        @{ n = 'Message'; e = { $_.Message } }

    $table | Format-Table -AutoSize -Wrap | Out-Host

    if ($display.Count -gt $shown.Count) {
        Write-Host ("  ... {0} more row(s) not shown. Use -ShowAll / -MaxDisplayRows 0, or open the CSV." -f ($display.Count - $shown.Count)) -ForegroundColor DarkGray
    }
    if (-not $ShowAll -and $PSCmdlet.ParameterSetName -ne 'Report') {
        $hidden = @($results | Where-Object { $_.Status -eq 'Skipped' }).Count
        if ($hidden -gt 0) {
            Write-Host ("  ({0} unchanged file(s) hidden - use -ShowAll to list them)" -f $hidden) -ForegroundColor DarkGray
        }
    }
}

try {
    $results | Export-Csv -LiteralPath $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host ''
    Write-Host "Report written to $ReportPath" -ForegroundColor Cyan
}
catch {
    Write-Warning "Could not write report to '$ReportPath': $($_.Exception.Message)"
}

Write-Host ''
Write-Host 'Summary' -ForegroundColor Green
$results | Group-Object Status | Sort-Object Name | ForEach-Object {
    Write-Host ('  {0,-10} {1}' -f $_.Name, $_.Count)
}

if ($PSCmdlet.ParameterSetName -eq 'Report') {
    Write-Host ''
    Write-Host 'Labels found' -ForegroundColor Green
    $results | Group-Object PreviousLabel | Sort-Object Count -Descending | ForEach-Object {
        $label = if ([string]::IsNullOrWhiteSpace($_.Name)) { '(no label)' } else { $_.Name }
        Write-Host ('  {0,-45} {1}' -f $label, $_.Count)
    }

    Write-Host ''
    Write-Host 'Label GUIDs (use one with -LabelId)' -ForegroundColor Green
    $results | Where-Object { $_.PreviousLabelId } |
        Group-Object PreviousLabelId |
        ForEach-Object { Write-Host ('  {0,-45} {1}' -f $_.Group[0].PreviousLabel, $_.Name) }
}

if ($PassThru) { $results }
