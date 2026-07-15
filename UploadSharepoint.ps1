# ===========================================
# Upload using Drive ID & Parent ID (Drive Item ID)
# ===========================================
$drive.id="b!lxV5xxxxxxxxxxxx-xxxxxxxxxxxxxxxxxx"
$folder.id="01EJVDTxxxxxxxxxxxxxxxxxxx"
$fileName="sample.txt"

$fileName = Split-Path $filePath -Leaf

$uploadUri = "https://graph.microsoft.com/v1.0/drives/{0}/items/{1}:/{2}:/content" -f `
    $drive.id, `
    $folder.id, `
    $fileName

$response = Invoke-MgGraphRequest `
    -Method PUT `
    -Uri $uploadUri `
    -InputFilePath $filePath `
    -ContentType "text/plain"

if ($response.id) {
    Write-Host ""
    Write-Host "Upload completed successfully." -ForegroundColor Green
    Write-Host "File Name : $($response.name)"
    Write-Host "File ID   : $($response.id)"
    Write-Host "Size      : $($response.size) bytes"
    Write-Host "Web URL   : $($response.webUrl)"
}
else {
    Write-Host "Upload failed. No response received." -ForegroundColor Red
}