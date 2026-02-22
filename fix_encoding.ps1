$path = 'D:\Cloud\roman@heimman.com\OneDrive - RH\Documents\Sharepoint Clean Tool\Sharepoint-Versions-Clean\CleanSharepointVersions.ps1'
$content = Get-Content $path -Raw -Encoding UTF8
$content | Out-File -FilePath $path -Encoding UTF8
Write-Host "File re-encoded to UTF8"
