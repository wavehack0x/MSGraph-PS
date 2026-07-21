# Requires:
# Install-Module Microsoft.Graph -Scope CurrentUser
# Install-Module ImportExcel -Scope CurrentUser


# ===========================================
# Prerequisite Checks
# ===========================================

if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {

    Write-Host "Microsoft Graph PowerShell SDK is not installed." -ForegroundColor Red
    Write-Host "Run: Install-Module Microsoft.Graph -Scope CurrentUser"
    Read-Host "Press Enter to exit"
    exit
}


if (-not (Get-Module -ListAvailable ImportExcel)) {

    Write-Host "ImportExcel module is not installed." -ForegroundColor Red
    Write-Host "Run: Install-Module ImportExcel -Scope CurrentUser"
    Read-Host "Press Enter to exit"
    exit
}


Import-Module Microsoft.Graph.Authentication
Import-Module ImportExcel



# ===========================================
# Connect Microsoft Graph
# ===========================================

try {

    $ctx = Get-MgContext

    if (-not $ctx) {
        throw
    }

}
catch {

    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow

    Connect-MgGraph `
        -Scopes "User.Read.All","AuditLog.Read.All"

    $ctx = Get-MgContext
}


Write-Host ""
Write-Host "Connected as: $($ctx.Account)" -ForegroundColor Green



# ===========================================
# Variables
# ===========================================

$inactiveCutoff = (Get-Date).AddDays(-180)
$pendingCutoff  = (Get-Date).AddDays(-90)
$recentCutoff   = (Get-Date).AddDays(-30)



# ===========================================
# Retrieve Guest Users
# ===========================================

Write-Host ""
Write-Host "Retrieving guest users..." -ForegroundColor Yellow


$guests = Get-MgUser `
    -Filter "userType eq 'Guest'" `
    -All `
    -Property `
    DisplayName,
    UserPrincipalName,
    Mail,
    CreatedDateTime,
    ExternalUserState,
    SignInActivity



Write-Host "Total guests found: $($guests.Count)" -ForegroundColor Green



# ===========================================
# Process Guests with Progress Bar
# ===========================================


$report = [System.Collections.Generic.List[Object]]::new()


$total = $guests.Count
$count = 0


foreach ($u in $guests) {


    $count++


    Write-Progress `
        -Activity "Processing Guest Users" `
        -Status "$($u.DisplayName) ($count/$total)" `
        -PercentComplete (($count / $total) * 100)



    $interactive = $u.SignInActivity.LastSignInDateTime

    $nonInteractive = $u.SignInActivity.LastNonInteractiveSignInDateTime



    $latest = @(
        $interactive,
        $nonInteractive
    ) |
    Where-Object { $_ } |
    Sort-Object -Descending |
    Select-Object -First 1



    $report.Add(
        [PSCustomObject]@{

            DisplayName = $u.DisplayName

            UserPrincipalName = $u.UserPrincipalName

            Mail = $u.Mail

            CreatedDateTime = $u.CreatedDateTime

            ExternalUserState = $u.ExternalUserState

            LastInteractiveSignIn = $interactive

            LastNonInteractiveSignIn = $nonInteractive

            LatestSignIn = $latest
        }
    )

}


Write-Progress `
    -Activity "Processing Guest Users" `
    -Completed



# ===========================================
# Create Reports
# ===========================================


$allGuests = $report



$inactiveGuests = $report | Where-Object {

    $_.ExternalUserState -eq "Accepted" -and
    $_.LatestSignIn -and
    $_.LatestSignIn -lt $inactiveCutoff

}



$acceptedNeverSignedIn = $report | Where-Object {

    $_.ExternalUserState -eq "Accepted" -and
    -not $_.LatestSignIn

}



$pendingGuests = $report | Where-Object {

    $_.ExternalUserState -eq "PendingAcceptance"

}



$pending90Guests = $report | Where-Object {

    $_.ExternalUserState -eq "PendingAcceptance" -and
    $_.CreatedDateTime -lt $pendingCutoff

}



$recentInvitations = $report | Where-Object {

    $_.CreatedDateTime -gt $recentCutoff

}



# ===========================================
# Export Excel with Progress
# ===========================================


$date = Get-Date -Format "yyyy-MM-dd"

$output = ".\GuestUserReview_$date.xlsx"



$exportTotal = 6
$exportCount = 0



function Export-Step {

    param(
        $Name
    )

    $script:exportCount++

    Write-Progress `
        -Activity "Exporting Excel Report" `
        -Status "$Name ($exportCount/$exportTotal)" `
        -PercentComplete (($exportCount/$exportTotal)*100)

}



Export-Step "All Guest Users"

$allGuests | Export-Excel $output `
    -WorksheetName "All Guest Users" `
    -AutoSize `
    -TableName AllGuests



Export-Step "Inactive >180 Days"

$inactiveGuests | Export-Excel $output `
    -WorksheetName "Inactive >180 Days" `
    -AutoSize `
    -TableName InactiveGuests `
    -Append



Export-Step "Accepted Never Signed In"

$acceptedNeverSignedIn | Export-Excel $output `
    -WorksheetName "Accepted Never Signed In" `
    -AutoSize `
    -TableName AcceptedNeverSignedIn `
    -Append



Export-Step "Pending Invitation"

$pendingGuests | Export-Excel $output `
    -WorksheetName "Pending Invitation" `
    -AutoSize `
    -TableName PendingInvitation `
    -Append



Export-Step "Pending >90 Days"

$pending90Guests | Export-Excel $output `
    -WorksheetName "Pending >90 Days" `
    -AutoSize `
    -TableName Pending90Days `
    -Append



Export-Step "Recent Invite <30 Days"

$recentInvitations | Export-Excel $output `
    -WorksheetName "Recent Invite <30 Days" `
    -AutoSize `
    -TableName RecentInvites `
    -Append



Write-Progress `
    -Activity "Exporting Excel Report" `
    -Completed



# ===========================================
# Summary
# ===========================================


Write-Host ""

Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "Guest User Review Summary" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan


Write-Host "Total Guest Users              : $($allGuests.Count)"
Write-Host "Inactive >180 Days             : $($inactiveGuests.Count)"
Write-Host "Accepted Never Signed In       : $($acceptedNeverSignedIn.Count)"
Write-Host "Pending Invitations            : $($pendingGuests.Count)"
Write-Host "Pending >90 Days               : $($pending90Guests.Count)"
Write-Host "Recent Invitations <30 Days    : $($recentInvitations.Count)"


Write-Host ""

Write-Host "Report generated successfully:" -ForegroundColor Green
Write-Host $output -ForegroundColor Yellow



Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "Press any key to exit..." -ForegroundColor Yellow
Write-Host "==============================================" -ForegroundColor Cyan


$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
