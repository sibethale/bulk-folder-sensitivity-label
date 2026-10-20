<#
.SYNOPSIS
    Lists the sensitivity labels in your tenant with their GUIDs, for use with
    Set-FolderSensitivityLabel.ps1 -LabelId.

.DESCRIPTION
    Connects to Security & Compliance PowerShell (Connect-IPPSSession from the
    ExchangeOnlineManagement module) and lists every label with its GUID, parent label and
    whether it is published. Runs in both Windows PowerShell 5.1 and PowerShell 7.

.PARAMETER UserPrincipalName
    UPN to sign in with. Omit to be prompted.

.PARAMETER IncludeUnpublished
    Also list labels that are not in any published label policy.

.EXAMPLE
    .\Get-SensitivityLabelList.ps1

.EXAMPLE
    .\Get-SensitivityLabelList.ps1 -UserPrincipalName admin@contoso.com -IncludeUnpublished
#>
[CmdletBinding()]
param(
    [string]$UserPrincipalName,
    [switch]$IncludeUnpublished
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw 'ExchangeOnlineManagement is not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop

$connected = $false
try {
    $null = Get-Command Get-Label -ErrorAction Stop
    $connected = $true
}
catch {
    $connectParams = @{ ErrorAction = 'Stop' }
    if ($UserPrincipalName) { $connectParams.UserPrincipalName = $UserPrincipalName }
    Write-Host 'Connecting to Security & Compliance PowerShell...' -ForegroundColor Cyan
    Connect-IPPSSession @connectParams
    $connected = $true
}

if (-not $connected) { throw 'Could not connect to Security & Compliance PowerShell.' }

$labels = Get-Label

if (-not $IncludeUnpublished) {
    $policies = @(Get-LabelPolicy)
    $publishedGuids = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]@($policies | ForEach-Object { $_.Labels } | ForEach-Object { [string]$_ }),
        [System.StringComparer]::OrdinalIgnoreCase
    )
    if ($publishedGuids.Count -gt 0) {
        $labels = $labels | Where-Object {
            $publishedGuids.Contains([string]$_.Guid) -or $publishedGuids.Contains([string]$_.Name) -or
            $publishedGuids.Contains([string]$_.DisplayName)
        }
    }
}

$labels |
    Sort-Object ParentId, DisplayName |
    Select-Object @{ n = 'Label'; e = { if ($_.ParentId) { "$($_.ParentLabelDisplayName) / $($_.DisplayName)" } else { $_.DisplayName } } },
                  @{ n = 'LabelId'; e = { $_.Guid } },
                  @{ n = 'Priority'; e = { $_.Priority } },
                  @{ n = 'Enabled'; e = { -not $_.Disabled } } |
    Format-Table -AutoSize

Write-Host ''
Write-Host 'Copy a LabelId into:' -ForegroundColor Green
Write-Host '  powershell.exe -NoProfile -File .\Set-FolderSensitivityLabel.ps1 -Path C:\Data -Recurse -LabelId <LabelId> -WhatIf'
