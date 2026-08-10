# ===========================================
# Get SharePoint Drive & Folder Information
# ===========================================
# Purpose:
# Retrieve SharePoint Site ID, Drive ID, and Folder ID
# using Microsoft Graph PowerShell
#
# Requirements:
# Install-Module Microsoft.Graph -Scope CurrentUser
#
# Permissions:
# Sites.ReadWrite.All
# Files.ReadWrite.All

# ===========================================
# Prerequisite Check
# ===========================================

if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Write-Host "Microsoft Graph PowerShell SDK is not installed." -ForegroundColor Red
    Write-Host "Install using:"
    Write-Host "Install-Module Microsoft.Graph -Scope CurrentUser"
    exit
}

# ===========================================
# Configuration
# ===========================================

$tenantHost = "contosso.sharepoint.com"
$sitePath   = "/sites/SampleSite"
$folderName = "SampleFolder" 


# ===========================================
# Connect to Microsoft Graph
# ===========================================

Connect-MgGraph -Scopes "Sites.ReadWrite.All","Files.ReadWrite.All"

$context = Get-MgContext

if (-not $context) {
    Write-Host "Microsoft Graph authentication failed." -ForegroundColor Red
    exit
}

Write-Host "Connected as: $($context.Account)" -ForegroundColor Green

# ===========================================
# Get SharePoint Site
# ===========================================

try {

    $siteUri = "https://graph.microsoft.com/v1.0/sites/{0}:{1}" -f `
        $tenantHost, `
        $sitePath

    $site = Invoke-MgGraphRequest `
        -Method GET `
        -Uri $siteUri

}
catch {

    Write-Host "Unable to retrieve SharePoint site." -ForegroundColor Red
    Write-Host $_.Exception.Message
    exit

}

# ===========================================
# Get Default Document Library
# ===========================================

$drive = Invoke-MgGraphRequest `
    -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drive"


# ===========================================
# Get Folder Information
# ===========================================

try {

    $folder = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root:/$folderName"

}
catch {

    Write-Host "Unable to find folder: $folderName" -ForegroundColor Red
    Write-Host $_.Exception.Message
    exit

}

# ===========================================
# Output Result
# ===========================================

Write-Host ""
Write-Host "SharePoint Information" -ForegroundColor Cyan
Write-Host "======================"
Write-Host "Site Name   : $($site.name)"
Write-Host "Site URL    : $($site.webUrl)"
Write-Host "Site ID     : $($site.id)"
Write-Host "Drive ID    : $($drive.id)"
Write-Host "Folder Name : $($folder.name)"
Write-Host "Folder ID   : $($folder.id)"

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "Completed successfully." -ForegroundColor Green
Write-Host "Press any key to exit..." -ForegroundColor Yellow
Write-Host "==============================================" -ForegroundColor Cyan

$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
