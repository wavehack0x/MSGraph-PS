# Requires:
# Install-Module ImportExcel -Scope CurrentUser
# Install-Module Microsoft.Graph.Users -Scope CurrentUser


# ===========================================
# Prerequisite Checks
# ===========================================

# Check Microsoft Graph Authentication module
if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    Write-Host "Microsoft Graph PowerShell SDK is not installed." -ForegroundColor Red
    Write-Host "Run the following command:"
    Write-Host "Install-Module Microsoft.Graph -Scope CurrentUser"
    return
}


# Check ImportExcel module
if (-not (Get-Module -ListAvailable ImportExcel)) {
    Write-Host "ImportExcel module is not installed." -ForegroundColor Red
    Write-Host "Run the following command:"
    Write-Host "Install-Module ImportExcel -Scope CurrentUser"
    return
}

# Import modules
Import-Module Microsoft.Graph.Authentication
Import-Module ImportExcel

# Verify Graph connection
try {
    $ctx = Get-MgContext

    if (-not $ctx) {
        throw
    }
}
catch {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes "User.Read.All","AuditLog.Read.All"
    $ctx = Get-MgContext
}

Write-Host "Connected as: $($ctx.Account)" -ForegroundColor Green
Write-Host "Prerequisite checks passed." -ForegroundColor Green

$cutoff = (Get-Date).AddDays(-180)

$guests = Get-MgUser `
    -Filter "userType eq 'Guest'" `
    -All `
    -Property DisplayName,UserPrincipalName,Mail,ExternalUserState,SignInActivity

$report = foreach ($u in $guests) {

    $interactive = $u.SignInActivity.LastSignInDateTime
    $nonInteractive = $u.SignInActivity.LastNonInteractiveSignInDateTime

    $latest = @($interactive,$nonInteractive) |
        Where-Object { $_ } |
        Sort-Object -Descending |
        Select-Object -First 1

    [PSCustomObject]@{
        DisplayName            = $u.DisplayName
        UserPrincipalName      = $u.UserPrincipalName
        Mail                   = $u.Mail
        ExternalUserState      = $u.ExternalUserState
        LastInteractiveSignIn  = $interactive
        LastNonInteractiveSignIn = $nonInteractive
        LatestSignIn           = $latest
    }
}

# Worksheet 1 - All Guests
$allGuests = $report

# Worksheet 2 - Inactive (>180 days)
$inactiveGuests = $report | Where-Object {
    $_.LatestSignIn -and $_.LatestSignIn -lt $cutoff
}

# Worksheet 3 - Pending Invitation / Never Signed In
$pendingGuests = $report | Where-Object {
    $_.ExternalUserState -eq "PendingAcceptance" -and
    -not $_.LatestSignIn
}

$date = Get-Date -Format "yyyy-MM-dd"
$output = ".\GuestUserReview_$date.xlsx"

$allGuests | Export-Excel $output `
    -WorksheetName "All Guest Users" `
    -AutoSize `
    -TableName AllGuests

$inactiveGuests | Export-Excel $output `
    -WorksheetName "Inactive >180 Days" `
    -AutoSize `
    -TableName InactiveGuests `
    -Append

$pendingGuests | Export-Excel $output `
    -WorksheetName "Pending Invitation" `
    -AutoSize `
    -TableName PendingGuests `
    -Append

Write-Host "Workbook exported to $output"
Write-Host "All Guests: $($allGuests.Count)"
Write-Host "Inactive (>180 days): $($inactiveGuests.Count)"
Write-Host "Pending Invitation: $($pendingGuests.Count)"

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "Report generation completed successfully." -ForegroundColor Green
Write-Host "Press any key to exit..." -ForegroundColor Yellow
Write-Host "==============================================" -ForegroundColor Cyan

$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")