$ErrorActionPreference = 'Stop'
$base = $PSScriptRoot
$script = Join-Path $base 'Set-FolderSensitivityLabel.ps1'

# syntax check
$errors = $null
$tokens = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile($script, [ref]$tokens, [ref]$errors)
if ($errors) { $errors | ForEach-Object { Write-Host "PARSE ERROR: $($_.Message)" }; exit 1 }
Write-Host 'Parse OK' -ForegroundColor Green

# sandbox tree
$sandbox = Join-Path $base 'sandbox'
if (Test-Path $sandbox) { Remove-Item $sandbox -Recurse -Force }
New-Item -ItemType Directory -Path (Join-Path $sandbox 'sub') | Out-Null
'a' | Set-Content (Join-Path $sandbox 'a.docx')
'b' | Set-Content (Join-Path $sandbox 'b.txt')
'c' | Set-Content (Join-Path $sandbox 'sub\c.docx')
'd' | Set-Content (Join-Path $sandbox '~$tmp.docx')

$target = 'bd0b5b6a-9d3b-4f2a-8b1d-1f1a2b3c4d5e'
$global:store = @{ (Join-Path $sandbox 'a.docx') = $target }  # already labelled
$global:setCalls = @()
$global:removeCalls = @()

function Get-FileStatus {
    param([string]$Path, [string]$ErrorAction)
    $id = $global:store[$Path]
    [pscustomobject]@{
        FileName    = $Path
        IsLabeled   = [bool]$id
        MainLabelId = $id
        MainLabelName = if ($id) { 'Confidential' } else { $null }
        SubLabelId  = $null
        SubLabelName = $null
    }
}

function Set-FileLabel {
    param([string]$Path, [string]$LabelId, [string]$JustificationMessage,
          [switch]$PreserveFileDetails, [string]$Owner, [string]$ErrorAction)
    $global:setCalls += $Path
    $global:store[$Path] = $LabelId
    [pscustomobject]@{ FileName = $Path; Status = 'Success' }
}

function Remove-FileLabel {
    param([string]$Path, [string]$JustificationMessage, [switch]$PreserveFileDetails, [string]$ErrorAction)
    $global:removeCalls += $Path
    $global:store.Remove($Path)
    [pscustomobject]@{ FileName = $Path; Status = 'Success' }
}

Write-Host "`n--- 1. Report mode ---" -ForegroundColor Yellow
$r = & $script -Path $sandbox -Recurse -ShowCurrentLabels -SkipModuleCheck -PassThru -ReportPath (Join-Path $base 'r1.csv')
if (@($r).Count -ne 4) { throw "expected 4 files, got $(@($r).Count)" }

Write-Host "`n--- 2. WhatIf apply with include/exclude ---" -ForegroundColor Yellow
$r = & $script -Path $sandbox -Recurse -LabelId $target -Include '*.docx' -Exclude '~$*' -SkipModuleCheck -PassThru -ReportPath (Join-Path $base 'r2.csv') -WhatIf
if (@($r).Count -ne 2) { throw "expected 2 filtered files, got $(@($r).Count)" }
if ($global:setCalls.Count -ne 0) { throw 'WhatIf must not call Set-FileLabel' }
$whatIf = @($r | Where-Object Status -eq 'WhatIf').Count
$skipped = @($r | Where-Object Message -eq 'Already has target label').Count
if ($whatIf -ne 1 -or $skipped -ne 1) { throw "expected 1 WhatIf + 1 already-labelled, got $whatIf/$skipped" }

Write-Host "`n--- 3. Real apply ---" -ForegroundColor Yellow
$r = & $script -Path $sandbox -Recurse -LabelId $target -Include '*.docx' -Exclude '~$*' -SkipModuleCheck -PassThru -ReportPath (Join-Path $base 'r3.csv') -Confirm:$false
if ($global:setCalls.Count -ne 1) { throw "expected 1 Set-FileLabel call, got $($global:setCalls.Count)" }
if ($global:store[(Join-Path $sandbox 'sub\c.docx')] -ne $target) { throw 'sub\c.docx not labelled' }

Write-Host "`n--- 4. Force relabel ---" -ForegroundColor Yellow
$global:setCalls = @()
$null = & $script -Path $sandbox -LabelId $target -Include '*.docx' -Exclude '~$*' -Force -SkipModuleCheck -PassThru -ReportPath (Join-Path $base 'r4.csv') -Confirm:$false
if ($global:setCalls.Count -ne 1) { throw "Force: expected 1 call, got $($global:setCalls.Count)" }

Write-Host "`n--- 5. Remove mode ---" -ForegroundColor Yellow
$null = & $script -Path $sandbox -Recurse -RemoveLabel -JustificationMessage 'declassified' -SkipModuleCheck -PassThru -ReportPath (Join-Path $base 'r5.csv') -Confirm:$false
if ($global:removeCalls.Count -ne 2) { throw "expected 2 removes, got $($global:removeCalls.Count)" }
if ($global:store.Count -ne 0) { throw 'labels remain after removal' }

Write-Host "`n--- 6. Missing folder error ---" -ForegroundColor Yellow
try { & $script -Path (Join-Path $base 'nope') -LabelId $target -SkipModuleCheck -PassThru -Confirm:$false; throw 'should have failed' }
catch { if ($_.Exception.Message -notlike 'Folder not found*') { throw } ; Write-Host "  ok: $($_.Exception.Message)" }

Write-Host "`n--- 7. Bad GUID rejected ---" -ForegroundColor Yellow
try { & $script -Path $sandbox -LabelId 'not-a-guid' -SkipModuleCheck -PassThru -Confirm:$false; throw 'should have failed' }
catch { Write-Host '  ok: validation rejected bad GUID' }

Write-Host "`nALL TESTS PASSED" -ForegroundColor Green
