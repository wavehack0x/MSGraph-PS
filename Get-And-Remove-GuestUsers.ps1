<#
===============================================================================
Script Name : Get-And-Remove-GuestUsers.ps1
Project     : Microsoft Entra ID Guest Lifecycle Automation
Description : Identifies and removes stale Microsoft Entra ID guest accounts
              based on configurable inactivity criteria while generating
              comprehensive logs and reports.

Author      : Waverly Chua
Created     : 2026-08-04
Version     : 1.0.0

Requirements:
    • PowerShell 7.0 or later
    • Microsoft Graph PowerShell SDK
    • Microsoft Graph Application Permissions
    • Global Administrator or appropriate Entra ID permissions

Copyright © 2026 Waverly Chua. All rights reserved.

This script is intended for internal use by authorized personnel only.
Unauthorized copying, redistribution, modification, or commercial use without
written permission from the author is prohibited.

THIS SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE, AND NON-INFRINGEMENT. IN NO EVENT SHALL THE
AUTHOR BE LIABLE FOR ANY CLAIM, DAMAGES, OR OTHER LIABILITY ARISING FROM THE
USE OF THIS SOFTWARE.

===============================================================================
Change Log
===============================================================================
v1.0.0  (2026-08-04)
- Initial production release.
- Combined guest user detail reporting and removal into a single script.
- Configurable inactivity threshold (-InactiveDays, default 180).
- Configurable pending-invitation threshold (-PendingDays, default 90).
- Criteria-based candidate selection (Accepted + inactive, OR Pending +
  stale invite) -- non-matching guests are excluded from both the report
  and removal.
- Optional scoping to a specific list of guests via -UserListPath.
- Report-only by default; no accounts are modified unless -Delete is
  explicitly passed.
- Interactive Y/N confirmation before deletion, skippable via -Force for
  unattended/automated runs.
- Per-guest txt detail report, including match reason and last sign-in.
- CSV export of processed accounts (Success / Failed / Skipped, with
  match reason and error message where applicable).
- Console summary counts (Removed / Failed / Skipped).
- Microsoft Graph authentication (delegated, interactive).
 
===============================================================================
#>

param(
    [string]$UserListPath=$null,
    [string]$DetailsPath=(Join-Path $PSScriptRoot "guest-details_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').txt"),
    [string]$ReportPath=(Join-Path $PSScriptRoot "GuestUserRemoval_Report_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').csv"),
    [int]$InactiveDays=180,
    [int]$PendingDays=90,
    [switch]$Delete,
    [switch]$Force
)

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Get-And-Remove-GuestUsers.ps1" -ForegroundColor White
Write-Host "Version : 1.0.0"
Write-Host "Author  : Waverly Chua"
Write-Host "Copyright $([char]0x00A9) 2026 Waverly Chua"
Write-Host "============================================================" -ForegroundColor Cyan

$inactiveCutoff = (Get-Date).AddDays(-$InactiveDays)
$pendingCutoff  = (Get-Date).AddDays(-$PendingDays)

function Exit-Script {
    Write-Host ""
    Write-Host "Press any key to exit..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# ===========================================
# Prerequisite Checks
# ===========================================

if (-not (Get-Module -ListAvailable Microsoft.Graph.Users)) {
    Write-Host "Microsoft Graph Users module is not installed." -ForegroundColor Red
    Write-Host "Install using: Install-Module Microsoft.Graph.Users -Scope CurrentUser"
    Exit-Script
}

Import-Module Microsoft.Graph.Users

# ===========================================
# Connect Microsoft Graph
# ===========================================

try {
    $ctx = Get-MgContext
    if (-not $ctx) {
        Connect-MgGraph -Scopes "User.ReadWrite.All", "AuditLog.Read.All"
        $ctx = Get-MgContext
    }
}
catch {
    Connect-MgGraph -Scopes "User.ReadWrite.All", "AuditLog.Read.All"
    $ctx = Get-MgContext
}

Write-Host ""
Write-Host "Connected as: $($ctx.Account)" -ForegroundColor Green

# ===========================================
# Retrieve Guest Users
# ===========================================

Write-Host ""
if ($UserListPath) {

    if (-not (Test-Path $UserListPath)) {
        Write-Host "User list file not found: $UserListPath" -ForegroundColor Red
        Exit-Script
    }

    Write-Host "Retrieving details for guests listed in: $UserListPath" -ForegroundColor Yellow
    $targetUpns = Get-Content $UserListPath | ForEach-Object { $_.Trim() } | Where-Object { $_ }

    $guests = foreach ($upn in $targetUpns) {
        try {
            $GUID=Get-MgUser -Filter "userPrincipalName eq '$UPN'"
            Get-MgUser -UserId $GUID.Id -Property Id,DisplayName,UserPrincipalName,Mail,UserType,CreatedDateTime,ExternalUserState,SignInActivity -ErrorAction Stop
        }
        catch {
            Write-Host "  Error retrieving user: $upn ($($_.Exception.Message))" -ForegroundColor Red
        }
    }

}
else {

    Write-Host "Retrieving all guest users in the tenant..." -ForegroundColor Yellow
    $guests = Get-MgUser -Filter "userType eq 'Guest'" -All `
        -Property Id,DisplayName,UserPrincipalName,Mail,UserType,CreatedDateTime,ExternalUserState,SignInActivity

}

if (-not $guests -or $guests.Count -eq 0) {
    Write-Host "No guest users found. Nothing to report or remove." -ForegroundColor Green
    Exit-Script
}

Write-Host "Total guest users found: $($guests.Count)" -ForegroundColor Cyan

# ===========================================
# Apply Inactive/Pending Criteria
# ===========================================

Write-Host ""
Write-Host "Evaluating against criteria: Inactive > $InactiveDays days OR Pending > $PendingDays days..." -ForegroundColor Yellow

$candidates = foreach ($g in $guests) {
    $interactive    = $g.SignInActivity.LastSignInDateTime
    $nonInteractive = $g.SignInActivity.LastNonInteractiveSignInDateTime
    $latest = @($interactive, $nonInteractive) | Where-Object { $_ } | Sort-Object -Descending | Select-Object -First 1

    $reason = $null
    if ($g.ExternalUserState -eq "Accepted" -and $latest -and $latest -lt $inactiveCutoff) {
        $reason = "Inactive > $InactiveDays days (last sign-in: $latest)"
    }
    elseif ($g.ExternalUserState -eq "PendingAcceptance" -and $g.CreatedDateTime -lt $pendingCutoff) {
        $reason = "Pending > $PendingDays days (invited: $($g.CreatedDateTime))"
    }

    if ($reason) {
        Add-Member -InputObject $g -MemberType NoteProperty -Name "MatchReason" -Value $reason -Force
        Add-Member -InputObject $g -MemberType NoteProperty -Name "LatestSignIn" -Value $latest -Force
        $g
    }
}

if (-not $candidates -or $candidates.Count -eq 0) {
    Write-Host "No guest users matched the inactive/pending criteria. Nothing to report or remove." -ForegroundColor Green
    Exit-Script
}

Write-Host "Guests matching criteria: $($candidates.Count) of $($guests.Count)" -ForegroundColor Cyan
$guests = $candidates

# ===========================================
# Write Details to TXT File
# ===========================================

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("Guest User Detail Report")
$lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$lines.Add("Criteria: Inactive > $InactiveDays days OR Pending > $PendingDays days")
$lines.Add("Total matching guests: $($guests.Count)")
$lines.Add("=" * 60)

foreach ($g in $guests) {
    $lines.Add("")
    $lines.Add("DisplayName        : $($g.DisplayName)")
    $lines.Add("UserPrincipalName  : $($g.UserPrincipalName)")
    $lines.Add("Mail               : $($g.Mail)")
    $lines.Add("Id                 : $($g.Id)")
    $lines.Add("UserType           : $($g.UserType)")
    $lines.Add("ExternalUserState  : $($g.ExternalUserState)")
    $lines.Add("CreatedDateTime    : $($g.CreatedDateTime)")
    $lines.Add("LatestSignIn       : $(if ($g.LatestSignIn) { $g.LatestSignIn } else { 'Never' })")
    $lines.Add("MatchReason        : $($g.MatchReason)")
    $lines.Add("-" * 60)
}

$lines | Out-File -FilePath $DetailsPath -Encoding UTF8

Write-Host ""
Write-Host "Guest details written to: $DetailsPath" -ForegroundColor Green

# ===========================================
# Report-Only Exit (default behavior)
# ===========================================

if (-not $Delete) {
    Write-Host ""
    Write-Host "[Report-Only Mode] No accounts were removed. Pass -Delete to remove these guests." -ForegroundColor Cyan
    Exit-Script
}

# ===========================================
# Confirm Deletion (unless -Force)
# ===========================================

if (-not $Force) {
    $confirm = Read-Host "`nProceed with deleting $($guests.Count) guest user(s)? (Y/N)"
    if ($confirm -ne "Y") {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        Exit-Script
    }
}
else {
    Write-Host ""
    Write-Host "-Force specified: skipping confirmation prompt." -ForegroundColor Yellow
}

# ===========================================
# Remove Guest Users
# ===========================================

$results = @()

foreach ($g in $guests) {

    Write-Host ""
    Write-Host "Removing: $($g.UserPrincipalName)" -ForegroundColor Yellow

    try {
        if ($g.UserType -ne "Guest") {
            Write-Host "  Skipped - Not a Guest account" -ForegroundColor DarkYellow
            $results += [PSCustomObject]@{
                UserPrincipalName = $g.UserPrincipalName
                DisplayName       = $g.DisplayName
                UserType          = $g.UserType
                MatchReason       = $g.MatchReason
                Status            = "Skipped"
                Message           = "User is not a Guest account"
                Date              = Get-Date
            }
            continue
        }

        Remove-MgUser -UserId $g.Id -Confirm:$false -ErrorAction Stop
        Write-Host "  Removed successfully" -ForegroundColor Green

        $results += [PSCustomObject]@{
            UserPrincipalName = $g.UserPrincipalName
            DisplayName       = $g.DisplayName
            UserType          = $g.UserType
            MatchReason       = $g.MatchReason
            Status            = "Success"
            Message           = "Guest user removed successfully"
            Date              = Get-Date
        }
    }
    catch {
        Write-Host "  Failed: $($_.Exception.Message)" -ForegroundColor Red
        $results += [PSCustomObject]@{
            UserPrincipalName = $g.UserPrincipalName
            DisplayName       = $g.DisplayName
            UserType          = $g.UserType
            MatchReason       = $g.MatchReason
            Status            = "Failed"
            Message           = $_.Exception.Message
            Date              = Get-Date
        }
    }
}

$results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

# ===========================================
# Summary
# ===========================================

$successCount = ($results | Where-Object { $_.Status -eq "Success" }).Count
$failedCount  = ($results | Where-Object { $_.Status -eq "Failed" }).Count
$skipCount    = ($results | Where-Object { $_.Status -eq "Skipped" }).Count

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "Guest User Removal Summary" -ForegroundColor Green
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "Total Processed : $($results.Count)"
Write-Host "Removed         : $successCount"
Write-Host "Failed          : $failedCount"
Write-Host "Skipped         : $skipCount"
Write-Host ""
Write-Host "Details file    : $DetailsPath"
Write-Host "Removal report  : $ReportPath"

if (-not $Force) {
    Exit-Script
}
