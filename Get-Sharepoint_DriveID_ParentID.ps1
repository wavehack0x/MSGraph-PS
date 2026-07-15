# ===========================================
# Configuration
# ===========================================
$tenantHost = "contosso.sharepoint.com"
$siteName   = "SampleSite"          # SharePoint Site Name
$folderName = "SampleFolder"               # Destination folder
$filePath   = "C:\Temp\test.txt"

# ===========================================
# Connect to Microsoft Graph
# ===========================================
Connect-MgGraph -Scopes "Sites.ReadWrite.All","Files.ReadWrite.All"

# ===========================================
# Get Site
# ===========================================
$siteUri = "https://graph.microsoft.com/v1.0/sites/{0}:/sites/{1}" -f $tenantHost, $siteName
$site = Invoke-MgGraphRequest -Method GET -Uri $siteUri

# ===========================================
# Get Default Document Library (Drive)
# ===========================================
$drive = Invoke-MgGraphRequest `
    -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drive"

Write-Host "Drive ID: $($drive.id)"

# ===========================================
# Get Folder (Parent ID)
# ===========================================
$folder = Invoke-MgGraphRequest `
    -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root:/$folderName"

Write-Host "Folder Name : $($folder.name)"
Write-Host "Parent ID   : $($folder.id)"

