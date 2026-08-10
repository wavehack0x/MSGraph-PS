# ===========================================
# Restore-DeletedGuestUsers.ps1
# ===========================================
#
# What it does:
#   Looks up soft-deleted guest users in Microsoft Entra ID by email
#   address (single or from a list), retrieves the Object ID Graph
#   needs for restoration, and restores them via
#   Restore-MgDirectoryDeletedItem.
#
# Why this is needed:
#   Restore-MgDirectoryDeletedItem requires the deleted object's GUID,
#   not the email/UPN. That GUID is only obtainable via
#   Get-MgDirectoryDeletedItemAsUser, which lists everything currently
#   in the recycle bin. This script does that lookup for you so you
#   don't have to hunt through the full deleted-items list by hand.
#
# Requirements:
#   Install-Module Microsoft.Graph.Users -Scope CurrentUser
#   Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
#
# Required Permissions:
#   User.ReadWrite.All  (list deleted users, restore them)
#
# Note: deleted users only remain recoverable for 30 days after removal.
# After that, Entra ID permanently deletes them and this script (or any
# tool) can no longer restore them.
#
# Usage:
#   # Restore a single guest by email (looks up the GUID for you)
#   .\Restore-DeletedGuestUsers.ps1 -Email "alice@external.com"
#
#   # Restore directly if you already know the Object ID
#   .\Restore-DeletedGuestUsers.ps1 -ObjectId "a9532b30-4edb-4b66-a3b0-6ac972a6065b"
#
#   # Bulk restore from a text file of emails, one per line
#   .\Restore-DeletedGuestUsers.ps1 -EmailListPath "C:\Temp\restore-list.txt"
#
#   # No parameters: interactive prompt
#   .\Restore-DeletedGuestUsers.ps1
#
#   # Skip the Y/N confirmation prompt (unattended use)
#   .\Restore-DeletedGuestUsers.ps1 -EmailListPath "C:\Temp\restore-list.txt" -Force
#
# ===========================================

param(
    [string]$Email          = $null,
    [string]$ObjectId       = $null,
    [string]$EmailListPath  = $null,
    [string]$ReportPath     = (Join-Path $PSScriptRoot "GuestRestore_Report_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').csv"),
    [switch]$Force
)

function Exit-Script {
    if (-not $Force) {
        Write-Host ""
        Write-Host "Press any key to exit..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    exit
}

function Confirm-YesNo {
    param([string]$Prompt)

    while ($true) {
        $response = Read-Host $Prompt
        switch ($response.Trim().ToUpper()) {
            "Y" { return $true }
            "N" { return $false }
            default {
                Write-Host "  Please type Y or N." -ForegroundColor Yellow
            }
        }
    }
}

function Select-FromMatches {
    param($MatchList)

    while ($true) {
        $choice = Read-Host "  Select which one to restore (number), or type S to skip"

        if ($choice.Trim().ToUpper() -eq "S") {
            return $null
        }

        $asInt = 0
        if ([int]::TryParse($choice.Trim(), [ref]$asInt) -and $asInt -ge 1 -and $asInt -le $MatchList.Count) {
            return $MatchList[$asInt - 1]
        }

        Write-Host "  Enter a number between 1 and $($MatchList.Count), or S to skip." -ForegroundColor Yellow
    }
}

function Read-Choice {
    # Loops until the user enters one of the allowed values (case-insensitive).
    param([string]$Prompt, [string[]]$AllowedValues)

    while ($true) {
        $response = (Read-Host $Prompt).Trim()
        if ($AllowedValues -contains $response) {
            return $response
        }
        Write-Host "  Please enter one of: $($AllowedValues -join ' / ')" -ForegroundColor Yellow
    }
}

function Read-RequiredEmail {
    # Loops until the user enters a non-empty value that looks like an
    # email address. Refuses blank input rather than silently accepting
    # it (an empty lookup value would otherwise match every deleted user).
    param([string]$Prompt)

    while ($true) {
        $response = (Read-Host $Prompt).Trim()
        if ($response -and $response -match '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            return $response
        }
        Write-Host "  Please enter a valid, non-empty email address." -ForegroundColor Yellow
    }
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
        Connect-MgGraph -Scopes "User.ReadWrite.All"
        $ctx = Get-MgContext
    }
}
catch {
    Connect-MgGraph -Scopes "User.ReadWrite.All"
    $ctx = Get-MgContext
}

Write-Host ""
Write-Host "Connected as: $($ctx.Account)" -ForegroundColor Green

# ===========================================
# Helper: find a deleted user's Object ID by email
# ===========================================

function Find-DeletedGuestByEmail {
    param([string]$LookupEmail, $DeletedUsersCache)

    # NOTE: On deletion, Entra ID can alter a user's UPN/mail (e.g.
    # appending a random string or timestamp) to free up the original
    # address for reuse. An exact or prefix match against the original
    # email can therefore miss the record entirely. Using a wildcard
    # "contains" match instead is more resilient -- the original email
    # fragment should still appear somewhere in the mangled value, even
    # if extra characters were added before/after it.

    $normalizedEmail = $LookupEmail.Trim().ToLower()

    # SAFETY: an empty/whitespace lookup value would turn into a "*"
    # wildcard below and match EVERY deleted user in the tenant. Refuse
    # it outright instead of silently matching everything.
    if (-not $normalizedEmail) {
        return @()
    }

    $localPart = ($normalizedEmail -split '@')[0]
    $mangledUpnFragment = $normalizedEmail.Replace('@', '_')

    $found = $DeletedUsersCache | Where-Object {
        ($_.Mail -and $_.Mail.ToLower() -like "*$normalizedEmail*") -or
        ($_.UserPrincipalName -and $_.UserPrincipalName.ToLower() -like "*$mangledUpnFragment*") -or
        ($_.UserPrincipalName -and $_.UserPrincipalName.ToLower() -like "*$localPart*")
    }

    return $found
}

# ===========================================
# Helper: restore a single deleted item by Object ID
# ===========================================

function Restore-SingleGuest {
    param([string]$Id, [string]$LabelForLog)

    try {
        Restore-MgDirectoryDeletedItem -DirectoryObjectId $Id -ErrorAction Stop
        Write-Host "  Restored successfully" -ForegroundColor Green
        return [PSCustomObject]@{
            Input   = $LabelForLog
            ObjectId = $Id
            Status  = "Success"
            Message = "Restored successfully"
            Date    = Get-Date
        }
    }
    catch {
        Write-Host "  Failed: $($_.Exception.Message)" -ForegroundColor Red
        return [PSCustomObject]@{
            Input   = $LabelForLog
            ObjectId = $Id
            Status  = "Failed"
            Message = $_.Exception.Message
            Date    = Get-Date
        }
    }
}

$results = @()

# ===========================================
# Path 1: Restore directly by known Object ID
# ===========================================

if ($ObjectId) {

    Write-Host ""
    Write-Host "Restoring by Object ID: $ObjectId" -ForegroundColor Yellow

    if (-not $Force) {
        if (-not (Confirm-YesNo "Proceed with restoring this object? (Y/N)")) {
            Write-Host "Cancelled." -ForegroundColor Yellow
            Exit-Script
        }
    }

    $results += Restore-SingleGuest -Id $ObjectId -LabelForLog $ObjectId
}

# ===========================================
# Path 1(a) : Restore by email (single or list) -- requires a lookup first
# ===========================================
else {

    # Determine the list of emails to process
    $emailsToProcess = @()

    if ($EmailListPath) {
        if (-not (Test-Path $EmailListPath)) {
            Write-Host "Email list file not found: $EmailListPath" -ForegroundColor Red
            Exit-Script
        }
        $emailsToProcess = Get-Content $EmailListPath | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }
    elseif ($Email) {
        $emailsToProcess = @($Email)
    }
    else {
        # Interactive fallback
        Write-Host ""
        Write-Host "No -Email, -ObjectId, or -EmailListPath provided."
        $mode = Read-Choice -Prompt "Restore (1) a single email or (2) a list from a file? [1/2]" -AllowedValues @("1", "2")
        if ($mode -eq "2") {
            $path = Read-Host "Path to email list file"
            if (-not (Test-Path $path)) {
                Write-Host "File not found: $path" -ForegroundColor Red
                Exit-Script
            }
            $emailsToProcess = Get-Content $path | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        }
        else {
            $singleEmail = Read-RequiredEmail -Prompt "Email address to restore"
            $emailsToProcess = @($singleEmail)
        }
    }

    if ($emailsToProcess.Count -eq 0) {
        Write-Host "No email addresses to process." -ForegroundColor Yellow
        Exit-Script
    }

    # Pull the full deleted-users list ONCE and search it in memory --
    # much faster than a separate Graph call per email.
    Write-Host ""
    Write-Host "Retrieving deleted user list from Microsoft Entra ID..." -ForegroundColor Yellow
    $deletedUsers = Get-MgDirectoryDeletedItemAsUser -All -Property Id,DisplayName,UserPrincipalName,Mail,DeletedDateTime

    Write-Host "Found $($deletedUsers.Count) deleted user(s) currently in the recycle bin." -ForegroundColor Cyan

    # Resolve each email to a matching deleted user
    $toRestore = @()
    foreach ($e in $emailsToProcess) {
        $foundMatches = Find-DeletedGuestByEmail -LookupEmail $e -DeletedUsersCache $deletedUsers

        if (-not $foundMatches -or $foundMatches.Count -eq 0) {
            Write-Host "  No deleted user found matching: $e" -ForegroundColor DarkYellow
            $results += [PSCustomObject]@{
                Input    = $e
                ObjectId = $null
                Status   = "NotFound"
                Message  = "No matching deleted user found (may be outside the 30-day recovery window, or already restored)"
                Date     = Get-Date
            }
            continue
        }

        if ($foundMatches.Count -gt 1) {
            Write-Host "  Multiple deleted users matched: $e" -ForegroundColor Yellow
            $i = 1
            foreach ($m in $foundMatches) {
                Write-Host "    [$i] $($m.DisplayName) | $($m.UserPrincipalName) | Deleted: $($m.DeletedDateTime) | Id: $($m.Id))"
                $i++
            }
            $selected = Select-FromMatches -MatchList $foundMatches
            if (-not $selected) {
                Write-Host "  Skipped." -ForegroundColor DarkYellow
                $results += [PSCustomObject]@{
                    Input    = $e
                    ObjectId = $null
                    Status   = "Skipped"
                    Message  = "Multiple matches, none selected"
                    Date     = Get-Date
                }
                continue
            }
        }
        else {
            $selected = $foundMatches
        }

        Write-Host "  Matched: $e -> $($selected.DisplayName) ($($selected.Id))" -ForegroundColor Green
        $toRestore += [PSCustomObject]@{ Email = $e; Match = $selected }
    }

    if ($toRestore.Count -eq 0) {
        Write-Host ""
        Write-Host "No matched deleted users to restore." -ForegroundColor Yellow
        $results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
        Write-Host "Report written to: $ReportPath"
        Exit-Script
    }

    Write-Host ""
    Write-Host "$($toRestore.Count) guest(s) matched and ready to restore:" -ForegroundColor Cyan
    foreach ($r in $toRestore) {
        Write-Host "  $($r.Email) -> $($r.Match.DisplayName) ($($r.Match.Id))"
    }

    if (-not $Force) {
        if (-not (Confirm-YesNo "`nProceed with restoring $($toRestore.Count) guest(s)? (Y/N)")) {
            Write-Host "Cancelled." -ForegroundColor Yellow
            Exit-Script
        }
    }

    foreach ($r in $toRestore) {
        Write-Host ""
        Write-Host "Restoring: $($r.Email) ($($r.Match.Id))" -ForegroundColor Yellow
        $results += Restore-SingleGuest -Id $r.Match.Id -LabelForLog $r.Email
    }
}

# ===========================================
# Report + Summary
# ===========================================

$results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

$successCount  = ($results | Where-Object { $_.Status -eq "Success" }).Count
$failedCount   = ($results | Where-Object { $_.Status -eq "Failed" }).Count
$notFoundCount = ($results | Where-Object { $_.Status -eq "NotFound" }).Count
$skipCount     = ($results | Where-Object { $_.Status -eq "Skipped" }).Count

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "Guest User Restore Summary" -ForegroundColor Green
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "Total Processed : $($results.Count)"
Write-Host "Restored        : $successCount"
Write-Host "Failed          : $failedCount"
Write-Host "Not Found       : $notFoundCount"
Write-Host "Skipped         : $skipCount"
Write-Host ""
Write-Host "Report          : $ReportPath"

if (-not $Force) {
    Exit-Script
}