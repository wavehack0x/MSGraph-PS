# MSGraph-PS
A collection of Microsoft Graph PowerShell scripts, automation projects, and learning examples for Microsoft 365, Entra ID, SharePoint, OneDrive, and Azure.

# Scripts
1. GetDriveId-ParentId.ps1
    Retrieves the SharePoint Drive ID and Parent Folder ID required for uploading files.

2. UploadToSharePoint.ps1
    Uploads a local file to a specified SharePoint folder using the Drive ID and Parent Folder ID obtained from the first script.

# Prerequisites
- Microsoft Graph PowerShell SDK
- SharePoint site access
- Microsoft Graph permissions:
    - Sites.ReadWrite.All
    - Files.ReadWrite.All

# How to Connect to Microsoft Graph
  Connect-MgGraph -Scopes "Sites.ReadWrite.All","Files.ReadWrite.All"

