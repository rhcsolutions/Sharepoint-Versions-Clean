<#
.SYNOPSIS
    SharePoint Version Cleanup Tool

.DESCRIPTION
    Bulk-deletes file version history from OneDrive personal sites across a Microsoft 365 tenant.
    Supports certificate authentication, saved config, and automatic app permission management.

.NOTES
    Requires: PnP PowerShell module (v3.x)
    Author:   RHC Solutions
    Web:      rhcsolutions.com
    Telegram: t.me/rhcsolutions
    Updated:  2026
#>

# ========== PARAMETERS ==========
param(
    [int]$MaxThreads = 10
)

# --- STARTUP ANIMATION + BANNER ---
function Show-Banner {
    $spinner = @('|', '/', '—', '\')
    for ($i = 0; $i -lt 16; $i++) {
        Write-Host "`r  $($spinner[$i % 4])  Initializing..." -NoNewline -ForegroundColor DarkCyan
        Start-Sleep -Milliseconds 60
    }
    Write-Host "`r                         " -NoNewline
    Write-Host ""
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║                                                    ║" -ForegroundColor Cyan
    Write-Host "  ║    SharePoint Version Cleanup Tool                  ║" -ForegroundColor White
    Write-Host "  ║                                                    ║" -ForegroundColor Cyan
    Write-Host "  ║    © RHC Solutions  •  rhcsolutions.com            ║" -ForegroundColor DarkCyan
    Write-Host "  ║    Telegram: t.me/rhcsolutions                     ║" -ForegroundColor DarkCyan
    Write-Host "  ║                                                    ║" -ForegroundColor Cyan
    Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}
Show-Banner

# ========== LOAD SAVED CONFIG ==========
$AppName = "SharePoint-Cleanup-Tool"
$CertPath = Join-Path $PSScriptRoot "SharePoint-Cleanup-Tool.pfx"
$ConfigPath = Join-Path $PSScriptRoot "config.json"
$AdminEmail = ""
$TenantName = ""
$ClientId = ""
$AppObjectId = ""

if (Test-Path $ConfigPath) {
    try {
        $Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        $AdminEmail = $Config.AdminEmail
        $TenantName = $Config.TenantName
        $ClientId = $Config.ClientId
    }
    catch {}
}

# ========== SAVED SETTINGS SHORTCUT ==========
if (-not [string]::IsNullOrWhiteSpace($AdminEmail) -and
    -not [string]::IsNullOrWhiteSpace($TenantName) -and
    -not [string]::IsNullOrWhiteSpace($ClientId)) {

    Write-Host "  Saved settings found:" -ForegroundColor Cyan
    Write-Host "    Admin:    $AdminEmail" -ForegroundColor White
    Write-Host "    Tenant:   $TenantName.onmicrosoft.com" -ForegroundColor White
    Write-Host "    App ID:   $ClientId" -ForegroundColor White
    Write-Host ""
    Write-Host "  [Enter] Use saved settings" -ForegroundColor White
    Write-Host "  [N]     Enter new settings (different tenant / new app)" -ForegroundColor White
    $SavedChoice = Read-Host "  Choice"
    if ($SavedChoice -eq 'N' -or $SavedChoice -eq 'n') {
        $AdminEmail = ""; $TenantName = ""; $ClientId = ""
    }
    Write-Host ""
}

# ========== STEP 1: ADMIN ACCOUNT ==========
if ([string]::IsNullOrWhiteSpace($AdminEmail)) {
    Write-Host "[STEP 1] Your SharePoint Administrator Account" -ForegroundColor Yellow
    Write-Host "─────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  • Must have the 'SharePoint Administrator' role" -ForegroundColor Gray
    Write-Host ""
    $AdminEmail = Read-Host "Admin account (e.g. admin@contoso.onmicrosoft.com)"
    if ([string]::IsNullOrWhiteSpace($AdminEmail)) {
        Write-Host "✗ Admin account is required. Exiting." -ForegroundColor Red
        exit
    }

    # Extract tenant name from email
    if ($AdminEmail -match '@(.+?)\.onmicrosoft\.com') {
        $TenantName = $matches[1]
    }
    elseif ($AdminEmail -match '@(.+?)\.') {
        $TenantName = $matches[1]
    }
    else {
        Write-Host "✗ Could not extract tenant name from email. Exiting." -ForegroundColor Red
        exit
    }

    Write-Host "✓ Admin account:  $AdminEmail" -ForegroundColor Green
    Write-Host "✓ Tenant:         $TenantName.onmicrosoft.com" -ForegroundColor Green
    Write-Host ""
}

# ========== STEP 2: APP REGISTRATION ==========
# Check certificate
if (Test-Path $CertPath) {
    Write-Host "  ✓ Certificate found: SharePoint-Cleanup-Tool.pfx" -ForegroundColor Green
}
else {
    Write-Host "  ⚠ Certificate not found at: $CertPath" -ForegroundColor Yellow
    Write-Host "    Will fall back to interactive browser login." -ForegroundColor Gray
}

if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
    Write-Host "  ✓ App ID: $ClientId" -ForegroundColor Cyan
    Write-Host "  [Enter] Use existing app   [N] Create new app" -ForegroundColor White
    $AppChoice = Read-Host "  Choice"
    if ($AppChoice -eq 'N' -or $AppChoice -eq 'n') { $ClientId = "" }
}

if ([string]::IsNullOrWhiteSpace($ClientId)) {
    Write-Host ""
    Write-Host "  Creating new app '$AppName'..." -ForegroundColor Yellow
    try {
        $AppRegistration = Register-PnPAzureADApp `
            -ApplicationName $AppName `
            -Tenant "$TenantName.onmicrosoft.com" `
            -SharePointDelegatePermissions "AllSites.FullControl" `
            -SharePointApplicationPermissions "Sites.FullControl.All" `
            -CertificatePath $CertPath `
            -ErrorAction Stop
        $ClientId = $AppRegistration.'AzureAppId/ClientId'
        Write-Host "  ✓ App created! Client ID: $ClientId" -ForegroundColor Green
        Write-Host ""
        Write-Host "  ⚠ Grant admin consent now (opening browser)..." -ForegroundColor Yellow
        Start-Process "https://login.microsoftonline.com/$TenantName.onmicrosoft.com/adminconsent?client_id=$ClientId"
        Start-Sleep -Seconds 2
        Write-Host "  Grant consent in the browser, then press any key to continue..." -ForegroundColor Cyan
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    catch {
        Write-Host "  ✗ Failed to create app: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Write-Host "  Paste an existing Client ID to continue, or press Ctrl+C to exit." -ForegroundColor Yellow
        Write-Host "  (Azure Portal → App registrations → $AppName → Application (client) ID)" -ForegroundColor Gray
        $ClientId = Read-Host "  Client ID"
        if ([string]::IsNullOrWhiteSpace($ClientId)) {
            Write-Host "✗ Client ID is required. Exiting." -ForegroundColor Red
            exit
        }
    }
}

# Save all settings for next run
@{ AdminEmail = $AdminEmail; TenantName = $TenantName; ClientId = $ClientId } |
ConvertTo-Json | Set-Content $ConfigPath -Encoding UTF8
Write-Host "✓ App ID: $ClientId" -ForegroundColor Green
Write-Host ""

# ========== DEFAULTS ==========
$VerboseLogging = $true   # Always verbose
$RemoveAccessAfter = $false  # Keep admin access after cleanup

# ========== STEP 3: TARGET USERS ==========
Write-Host "[STEP 3] Target Users (whose OneDrive versions will be deleted)" -ForegroundColor Yellow
Write-Host "─────────────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  [A] ALL users in the tenant" -ForegroundColor White
Write-Host "  [S] SPECIFIC user(s) — you will pick from a list" -ForegroundColor White
Write-Host ""
$ScopeChoice = Read-Host "Choice (A/S)"
Write-Host ""

# ========== CONFIGURATION SUMMARY ==========
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "                CONFIGURATION SUMMARY                     " -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Tenant:         $TenantName.onmicrosoft.com" -ForegroundColor White
Write-Host "  Admin account:  $AdminEmail" -ForegroundColor Cyan
Write-Host "  App ID:         $ClientId" -ForegroundColor White
Write-Host "  Target users:   $(if ($ScopeChoice -eq 'A') { 'ALL users in tenant' } else { 'SPECIFIC users (selected after discovery)' })" -ForegroundColor Green
Write-Host "  Logging:        VERBOSE (default)" -ForegroundColor White
Write-Host "  Remove access:  NO — admin stays as site owner (default)" -ForegroundColor White
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "Press any key to connect and start discovery, or Ctrl+C to cancel..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Write-Host ""

# ========== BUILD URLS ==========
$AdminUrl = "https://$TenantName-admin.sharepoint.com"
$MySiteHost = "https://$TenantName-my.sharepoint.com"

# ========== CONNECT TO ADMIN CENTER ==========
Write-Host "[CONNECTION] Connecting to SharePoint Admin Center..." -ForegroundColor Yellow
Write-Host "  URL: $AdminUrl" -ForegroundColor Gray

try {
    if (Test-Path $CertPath) {
        Write-Host "  Using certificate authentication (no browser needed)" -ForegroundColor Gray
        Connect-PnPOnline `
            -Url $AdminUrl `
            -ClientId $ClientId `
            -Tenant "$TenantName.onmicrosoft.com" `
            -CertificatePath $CertPath `
            -ErrorAction Stop
    }
    else {
        Write-Host "  No certificate — opening browser for $AdminEmail" -ForegroundColor Gray
        Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId -ErrorAction Stop
    }
    Write-Host "✓ Connected!" -ForegroundColor Green
}
catch {
    Write-Host "✗ Connection failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "  • Verify the Client ID is correct for app '$AppName'" -ForegroundColor Yellow
    Write-Host "  • Ensure the certificate is uploaded to the app in Azure Portal" -ForegroundColor Yellow
    Write-Host "  • Admin consent must be granted for Sites.FullControl.All" -ForegroundColor Yellow
    exit
}

# ========== ENSURE APP PERMISSIONS ==========
# Look up the app's Object ID via Graph and ensure Sites.FullControl.All is granted
Write-Host "[PERMISSIONS] Verifying app permissions..." -ForegroundColor Yellow
try {
    $GraphToken = Get-PnPGraphAccessToken
    $AppInfo = Invoke-RestMethod `
        -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=appId eq '$ClientId'" `
        -Headers @{ Authorization = "Bearer $GraphToken" } `
        -ErrorAction Stop
    $AppObjectId = $AppInfo.value[0].id

    $SharePointAppId = "00000003-0000-0ff1-ce00-000000000000"
    $AllSitesFullControl = "56680e0d-d2a3-4ae5-9b02-e4a12ea7f3b9"  # Delegated
    $SitesFullControlAll = "678536fe-1083-478a-9c59-b99265e6b0d3"  # Application

    $Body = @{
        requiredResourceAccess = @(
            @{
                resourceAppId  = $SharePointAppId
                resourceAccess = @(
                    @{ id = $AllSitesFullControl; type = "Scope" }
                    @{ id = $SitesFullControlAll; type = "Role" }
                )
            }
        )
    } | ConvertTo-Json -Depth 5

    Invoke-RestMethod `
        -Method Patch `
        -Uri "https://graph.microsoft.com/v1.0/applications/$AppObjectId" `
        -Headers @{ Authorization = "Bearer $GraphToken"; "Content-Type" = "application/json" } `
        -Body $Body `
        -ErrorAction Stop

    Write-Host "✓ Permissions confirmed: AllSites.FullControl (Delegated) + Sites.FullControl.All (Application)" -ForegroundColor Green
}
catch {
    Write-Host "⚠ Could not auto-verify permissions: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "  If discovery fails, manually add in Azure Portal:" -ForegroundColor Gray
    Write-Host "  App registrations → $AppName → API permissions → SharePoint → Sites.FullControl.All (Application)" -ForegroundColor Gray
    Write-Host "  Then grant admin consent." -ForegroundColor Gray
}
Write-Host ""

# ========== RETRIEVE SharePoint SITES ==========
Write-Host "[DISCOVERY] Scanning for OneDrive sites..." -ForegroundColor Yellow
Write-Host "(This may take a few minutes depending on tenant size)" -ForegroundColor Gray

try {
    $AllSites = Get-PnPTenantSite -IncludeOneDriveSites -ErrorAction Stop
}
catch {
    Write-Host "✗ Failed to retrieve tenant sites!" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Possible causes:" -ForegroundColor Yellow
    Write-Host "  1. The app '$AppName' is missing the Sites.FullControl.All Application permission" -ForegroundColor Gray
    Write-Host "     → Azure Portal → App registrations → $AppName → API permissions" -ForegroundColor Gray
    Write-Host "     → Add SharePoint → Application → Sites.FullControl.All → Grant admin consent" -ForegroundColor Gray
    Write-Host "  2. The certificate is not uploaded to the app" -ForegroundColor Gray
    Write-Host "     → Azure Portal → App registrations → $AppName → Certificates & secrets" -ForegroundColor Gray
    Write-Host "     → Upload SharePoint-Cleanup-Tool.cer" -ForegroundColor Gray
    exit
}

$SharePointSites = $AllSites | Where-Object { $_.Url -like "$MySiteHost/personal/*" }
$Count = $SharePointSites.Count
Write-Host "✓ Found $Count OneDrive site(s)" -ForegroundColor Green

if ($Count -eq 0) {
    Write-Host "✗ No OneDrive sites found." -ForegroundColor Red
    Write-Host "  Retrieved $($AllSites.Count) total site(s) but none matched: $MySiteHost/personal/*" -ForegroundColor Yellow
    exit
}
Write-Host ""

# ========== USER SELECTION ==========
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "   TARGET USERS — OneDrive owners found in tenant ($Count)  " -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan
$i = 0
foreach ($Site in $SharePointSites) {
    $i++
    Write-Host "  [$i] $($Site.Owner)" -ForegroundColor Green
}
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$TargetSites = @()

if ($ScopeChoice -eq "S") {
    Write-Host "Select target user(s) by number (whose versions will be deleted):" -ForegroundColor Yellow
    Write-Host "(Separate multiple numbers with commas, e.g., '1,3,5' or single number '2')" -ForegroundColor Gray
    $SelectedNumbers = Read-Host "User number(s)"
    
    # Parse the input numbers
    $NumberArray = $SelectedNumbers -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }
    
    if ($NumberArray.Count -eq 0) {
        Write-Host "✗ No valid numbers entered. Exiting." -ForegroundColor Red
        exit
    }
    
    # Validate and collect selected sites
    $SelectedEmails = @()
    foreach ($num in $NumberArray) {
        if ($num -ge 1 -and $num -le $SharePointSites.Count) {
            $SelectedEmails += $SharePointSites[$num - 1].Owner
        }
        else {
            Write-Host "⚠ Warning: User number $num is out of range (1-$($SharePointSites.Count)) - skipping" -ForegroundColor Yellow
        }
    }
    
    if ($SelectedEmails.Count -eq 0) {
        Write-Host "✗ No valid users selected. Exiting." -ForegroundColor Red
        exit
    }
    
    $TargetSites = $SharePointSites | Where-Object { $SelectedEmails -contains $_.Owner }
    
    Write-Host ""
    Write-Host "Selected user(s):" -ForegroundColor Cyan
    foreach ($site in $TargetSites) {
        Write-Host "  ✓ $($site.Owner)" -ForegroundColor Green
    }
    Write-Host ""
    Write-Host "✓ Targeting $($TargetSites.Count) user(s)" -ForegroundColor Green
}
else {
    $TargetSites = $SharePointSites
    Write-Host "✓ Targeting ALL $Count user(s)" -ForegroundColor Green
}
Write-Host ""

# ========== FINAL CONFIRMATION ==========
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Red
Write-Host "                    ⚠ WARNING ⚠                          " -ForegroundColor Red
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Red
Write-Host "About to DELETE ALL VERSION HISTORY for:" -ForegroundColor Yellow
Write-Host "  • Users: $($TargetSites.Count)" -ForegroundColor White
Write-Host "  • This action CANNOT be undone!" -ForegroundColor Red
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Red
Write-Host ""
$Confirm = Read-Host "Type 'DELETE' to confirm (or anything else to cancel)"

if ($Confirm -ne "DELETE") {
    Write-Host "✗ Operation cancelled by user." -ForegroundColor Yellow
    exit
}
Write-Host ""

# ========== PROCESSING LOOP ==========
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "              STARTING VERSION CLEANUP                    " -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "   Parallel processing: $MaxThreads threads" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

function Process-UserSite {
    param(
        [string]$Url,
        [string]$Owner,
        [string]$AdminEmail,
        [string]$ClientId,
        [string]$CertPath,
        [int]$ThreadId
    )
    
    $result = [PSCustomObject]@{
        ThreadId         = $ThreadId
        Owner            = $Owner
        Url              = $Url
        Success          = $false
        FilesProcessed   = 0
        VersionsDeleted  = 0
        FilesSkipped     = 0
        Errors           = 0
        ErrorMessage     = ""
    }
    
    try {
        $UserConn = $null
        
        try {
            Set-PnPTenantSite -Url $Url -Owners @($AdminEmail) -ErrorAction SilentlyContinue | Out-Null
            Start-Sleep -Seconds 2
        }
        catch { }
        
        $UserConn = Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId -ReturnConnection -ErrorAction Stop
        
        $DocumentLibrary = $null
        $PossibleNames = @("Documents", "Shared Documents", "Documenten", "Dokumenty", "Documentos")
        
        foreach ($LibName in $PossibleNames) {
            try {
                $TestLib = Get-PnPList -Identity $LibName -Connection $UserConn -ErrorAction SilentlyContinue
                if ($TestLib) {
                    $DocumentLibrary = $TestLib
                    break
                }
            }
            catch { }
        }
        
        if (-not $DocumentLibrary) {
            $AllLists = Get-PnPList -Connection $UserConn
            $DocumentLibrary = $AllLists | Where-Object { 
                $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false 
            } | Select-Object -First 1
        }
        
        if (-not $DocumentLibrary) {
            $result.ErrorMessage = "No document library found"
            return $result
        }
        
        $LibraryName = $DocumentLibrary.Title
        
        try {
            Set-PnPList -Identity $LibraryName -EnableVersioning $true -Connection $UserConn -ErrorAction SilentlyContinue | Out-Null
            Set-PnPList -Identity $LibraryName -MajorVersions 1 -Connection $UserConn -ErrorAction SilentlyContinue | Out-Null
        }
        catch { }
        
        $AllFiles = @()
        try {
            $AllFiles = Get-PnPListItem -List $LibraryName -PageSize 1000 -Connection $UserConn -Fields "FileLeafRef", "FileRef", "File_x0020_Size", "Modified" | 
                Where-Object { $_.FileSystemObjectType -eq "File" }
        }
        catch {
            try {
                $AllFiles = Get-PnPListItem -List $LibraryName -PageSize 500 -Connection $UserConn | 
                    Where-Object { $_.FileSystemObjectType -eq "File" }
            }
            catch {
                $result.ErrorMessage = "Failed to scan library"
                return $result
            }
        }
        
        $TotalFileCount = $AllFiles.Count
        
        if ($TotalFileCount -eq 0) {
            $result.Success = $true
            $result.FilesProcessed = 0
            return $result
        }
        
        $UserFilesWithVersions = 0
        $UserFilesSkipped = 0
        $ErrorCount = 0
        $TotalVersionsCount = 0
        
        foreach ($item in $AllFiles) {
            $fileName = $item.FieldValues.FileLeafRef
            $fileUrl = $item.FieldValues.FileRef
            
            try {
                $versions = Get-PnPFileVersion -Url $fileUrl -Connection $UserConn -ErrorAction SilentlyContinue
                $versionCount = if ($versions) { $versions.Count } else { 0 }
                
                if ($versionCount -gt 0) {
                    $retryCount = 0
                    $maxRetries = 3
                    $success = $false
                    
                    while (-not $success -and $retryCount -lt $maxRetries) {
                        try {
                            Remove-PnPFileVersion -Url $fileUrl -All -Force -Connection $UserConn -ErrorAction Stop
                            $success = $true
                            $UserFilesWithVersions++
                            $TotalVersionsCount += $versionCount
                        }
                        catch {
                            $errorMsg = $_.Exception.Message
                            if ($errorMsg -like "*No file versions*" -or $errorMsg -like "*versions to delete*" -or $errorMsg -like "*not found*") {
                                $success = $true
                            }
                            else {
                                $retryCount++
                                if ($retryCount -lt $maxRetries) {
                                    Start-Sleep -Milliseconds 500
                                }
                            }
                        }
                    }
                    
                    if (-not $success) {
                        $ErrorCount++
                    }
                }
                else {
                    $UserFilesSkipped++
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                if ($errorMsg -notlike "*No file versions*" -and $errorMsg -notlike "*versions to delete*") {
                    $ErrorCount++
                }
                else {
                    $UserFilesSkipped++
                }
            }
        }
        
        $result.Success = $true
        $result.FilesProcessed = $TotalFileCount
        $result.VersionsDeleted = $TotalVersionsCount
        $result.FilesSkipped = $UserFilesSkipped
        $result.Errors = $ErrorCount
    }
    catch {
        $result.ErrorMessage = $_.Exception.Message
    }
    
    return $result
}

$TotalFilesProcessed = 0
$TotalVersionsDeleted = 0
$FailedSites = @()
$CompletedSites = @()
$UserNumber = 0

$runspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
$runspacePool.Open()

$jobs = @()

Write-Host "[PROCESSING] Starting parallel processing..." -ForegroundColor Yellow
Write-Host ""

foreach ($Site in $TargetSites) {
    $UserNumber++
    $Url = $Site.Url
    $Owner = $Site.Owner
    
    Write-Host "[$UserNumber/$($TargetSites.Count)] Queuing: $Owner" -ForegroundColor Cyan
    
    $powershell = [powershell]::Create().AddScript({
        param($Url, $Owner, $AdminEmail, $ClientId, $CertPath, $ThreadId)
        Process-UserSite -Url $Url -Owner $Owner -AdminEmail $AdminEmail -ClientId $ClientId -CertPath $CertPath -ThreadId $ThreadId
    }).AddParameter("Url", $Url).AddParameter("Owner", $Owner).AddParameter("AdminEmail", $AdminEmail).AddParameter("ClientId", $ClientId).AddParameter("CertPath", $CertPath).AddParameter("ThreadId", $UserNumber)
    
    $powershell.RunspacePool = $runspacePool
    
    $jobs += [PSCustomObject]@{
        PowerShell = $powershell
        Handle     = $powershell.BeginInvoke()
        Owner      = $Owner
        Url        = $Url
    }
}

Write-Host ""
Write-Host "[PROCESSING] Waiting for all jobs to complete..." -ForegroundColor Yellow
Write-Host ""

$jobResults = @()
$completed = 0

while ($jobs.Count -gt 0) {
    foreach ($job in @($jobs)) {
        if ($job.Handle.IsCompleted) {
            try {
                $result = $job.PowerShell.EndInvoke($job.Handle)
                $jobResults += $result
                
                $completed++
                
                if ($result.Success) {
                    Write-Host "[$completed/$($TargetSites.Count)] OK $Owner - $($result.FilesProcessed) files, $($result.VersionsDeleted) versions deleted" -ForegroundColor Green
                    $TotalFilesProcessed += $result.FilesProcessed
                    $TotalVersionsDeleted += $result.VersionsDeleted
                    $CompletedSites += [PSCustomObject]@{
                        Owner            = $result.Owner
                        Url              = $result.Url
                        FilesProcessed   = $result.FilesProcessed
                        VersionsDeleted  = $result.VersionsDeleted
                        FilesSkipped     = $result.FilesSkipped
                        Errors           = $result.Errors
                    }
                }
                else {
                    Write-Host "[$completed/$($TargetSites.Count)] FAIL $Owner - $($result.ErrorMessage)" -ForegroundColor Red
                    $FailedSites += [PSCustomObject]@{
                        Owner = $result.Owner
                        Url   = $result.Url
                        Error = $result.ErrorMessage
                    }
                }
            }
            catch {
                Write-Host "[$completed/$($TargetSites.Count)] ERROR $Owner - $($_.Exception.Message)" -ForegroundColor Red
                $FailedSites += [PSCustomObject]@{
                    Owner = $job.Owner
                    Url   = $job.Url
                    Error = $_.Exception.Message
                }
            }
            
            $job.PowerShell.Dispose()
            $jobs.Remove($job)
        }
    }
    
    if ($jobs.Count -gt 0) {
        Start-Sleep -Milliseconds 500
    }
}

$runspacePool.Close()
$runspacePool.Dispose()

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "                   PROCESSING COMPLETE                      " -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

if ($CompletedSites.Count -gt 0) {
    Write-Host "┌─────────────────────────────────────────────────────────┐" -ForegroundColor Green
    Write-Host "│                  SUCCESSFUL SITES                       │" -ForegroundColor Green
    Write-Host "└─────────────────────────────────────────────────────────┘" -ForegroundColor Green
    
    foreach ($site in $CompletedSites) {
        Write-Host "  User:      $($site.Owner)" -ForegroundColor White
        Write-Host "  Files:     $($site.FilesProcessed)" -ForegroundColor Gray
        Write-Host "  Versions:  $($site.VersionsDeleted)" -ForegroundColor Gray
        Write-Host "  Skipped:   $($site.FilesSkipped)" -ForegroundColor Gray
        Write-Host "  Errors:    $($site.Errors)" -ForegroundColor Gray
        Write-Host "  ----------------------------------------" -ForegroundColor DarkGray
    }
    
    Write-Host ""
    Write-Host "  TOTAL FILES PROCESSED:    $TotalFilesProcessed" -ForegroundColor Cyan
    Write-Host "  TOTAL VERSIONS DELETED:   $TotalVersionsDeleted" -ForegroundColor Cyan
    Write-Host ""
}

# ========== FINAL SUMMARY ==========
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "                   ✓ JOB COMPLETE ✓                      " -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "Users Targeted:      $($TargetSites.Count)" -ForegroundColor White
Write-Host "Files Processed:     $TotalFilesProcessed" -ForegroundColor White
Write-Host "Versions Deleted:    $TotalVersionsDeleted" -ForegroundColor Cyan
Write-Host "Failed Sites:        $($FailedSites.Count)" -ForegroundColor $(if ($FailedSites.Count -gt 0) { "Red" } else { "Green" })
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green

if ($FailedSites.Count -gt 0) {
    Write-Host ""
    Write-Host "┌──────────────────────────────────────────────────────┐" -ForegroundColor Red
    Write-Host "│                  FAILED SITES                        │" -ForegroundColor Red
    Write-Host "└──────────────────────────────────────────────────────┘" -ForegroundColor Red
    foreach ($Failed in $FailedSites) {
        Write-Host "  User:  $($Failed.Owner)" -ForegroundColor Yellow
        Write-Host "  Error: $($Failed.Error)" -ForegroundColor Red
        Write-Host "  ───────────────────────────────────────" -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "Script completed at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan



