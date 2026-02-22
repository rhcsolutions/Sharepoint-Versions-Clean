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
        Write-Host "`r  $($spinner[$i % 4])  Initializing..." -NoNewline -ForegroundColor Green
        Start-Sleep -Milliseconds 60
    }
    Write-Host "`r                         " -NoNewline
    Write-Host ""
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "  ║                                                    ║" -ForegroundColor Green
    Write-Host "  ║    SharePoint Version Cleanup Tool                  ║" -ForegroundColor Green
    Write-Host "  ║                                                    ║" -ForegroundColor Green
    Write-Host "  ║    © RHC Solutions  •  rhcsolutions.com            ║" -ForegroundColor Green
    Write-Host "  ║    Telegram: t.me/rhcsolutions                     ║" -ForegroundColor Green
    Write-Host "  ║                                                    ║" -ForegroundColor Green
    Write-Host "  ╚══════════════════════════════════════════════════════╝" -ForegroundColor Green
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

    Write-Host "  Saved settings found:" -ForegroundColor Green
    Write-Host "    Admin:    $AdminEmail" -ForegroundColor Green
    Write-Host "    Tenant:   $TenantName.onmicrosoft.com" -ForegroundColor Green
    Write-Host "    App ID:   $ClientId" -ForegroundColor Green
    Write-Host ""
    Write-Host "  [Enter] Use saved settings" -ForegroundColor Green
    Write-Host "  [N]     Enter new settings (different tenant / new app)" -ForegroundColor Green
    $SavedChoice = Read-Host "  Choice"
    if ($SavedChoice -eq 'N' -or $SavedChoice -eq 'n') {
        $AdminEmail = ""; $TenantName = ""; $ClientId = ""
    }
    Write-Host ""
}

# ========== STEP 1: ADMIN ACCOUNT ==========
if ([string]::IsNullOrWhiteSpace($AdminEmail)) {
    Write-Host "[STEP 1] Your SharePoint Administrator Account" -ForegroundColor Green
    Write-Host "─────────────────────────────────────────────────────────" -ForegroundColor Green
    Write-Host "  • Must have the 'SharePoint Administrator' role" -ForegroundColor Green
    Write-Host ""
    $AdminEmail = Read-Host "Admin account (e.g. admin@contoso.onmicrosoft.com)"
    if ([string]::IsNullOrWhiteSpace($AdminEmail)) {
        Write-Host "✗ Admin account is required. Exiting." -ForegroundColor Green
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
        Write-Host "✗ Could not extract tenant name from email. Exiting." -ForegroundColor Green
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
    Write-Host "  ⚠ Certificate not found at: $CertPath" -ForegroundColor Green
    Write-Host "    Will fall back to interactive browser login." -ForegroundColor Green
}

if (-not [string]::IsNullOrWhiteSpace($ClientId)) {
    Write-Host "  ✓ App ID: $ClientId" -ForegroundColor Green
    Write-Host "  [Enter] Use existing app   [N] Create new app" -ForegroundColor Green
    $AppChoice = Read-Host "  Choice"
    if ($AppChoice -eq 'N' -or $AppChoice -eq 'n') { $ClientId = "" }
}

if ([string]::IsNullOrWhiteSpace($ClientId)) {
    Write-Host ""
    Write-Host "  Creating new app '$AppName'..." -ForegroundColor Green
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
        Write-Host "  ⚠ Grant admin consent now (opening browser)..." -ForegroundColor Green
        Start-Process "https://login.microsoftonline.com/$TenantName.onmicrosoft.com/adminconsent?client_id=$ClientId"
        Start-Sleep -Seconds 2
        Write-Host "  Grant consent in the browser, then press any key to continue..." -ForegroundColor Green
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
    catch {
        Write-Host "  ✗ Failed to create app: $($_.Exception.Message)" -ForegroundColor Green
        Write-Host ""
        Write-Host "  Paste an existing Client ID to continue, or press Ctrl+C to exit." -ForegroundColor Green
        Write-Host "  (Azure Portal → App registrations → $AppName → Application (client) ID)" -ForegroundColor Green
        $ClientId = Read-Host "  Client ID"
        if ([string]::IsNullOrWhiteSpace($ClientId)) {
            Write-Host "✗ Client ID is required. Exiting." -ForegroundColor Green
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
Write-Host "[STEP 3] Target Users (whose OneDrive versions will be deleted)" -ForegroundColor Green
Write-Host "─────────────────────────────────────────────────────────" -ForegroundColor Green
Write-Host "  [A] ALL users in the tenant" -ForegroundColor Green
Write-Host "  [S] SPECIFIC user(s) — you will pick from a list" -ForegroundColor Green
Write-Host ""
$ScopeChoice = Read-Host "Choice (A/S)"
Write-Host ""

# ========== CONFIGURATION SUMMARY ==========
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "                CONFIGURATION SUMMARY                     " -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  Tenant:         $TenantName.onmicrosoft.com" -ForegroundColor Green
Write-Host "  Admin account:  $AdminEmail" -ForegroundColor Green
Write-Host "  App ID:         $ClientId" -ForegroundColor Green
Write-Host "  Target users:   $(if ($ScopeChoice -eq 'A') { 'ALL users in tenant' } else { 'SPECIFIC users (selected after discovery)' })" -ForegroundColor Green
Write-Host "  Logging:        VERBOSE (default)" -ForegroundColor Green
Write-Host "  Remove access:  NO — admin stays as site owner (default)" -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""
Write-Host "Press any key to connect and start discovery, or Ctrl+C to cancel..." -ForegroundColor Green
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Write-Host ""

# ========== BUILD URLS ==========
$AdminUrl = "https://$TenantName-admin.sharepoint.com"
$MySiteHost = "https://$TenantName-my.sharepoint.com"

# ========== CONNECT TO ADMIN CENTER ==========
Write-Host "[CONNECTION] Connecting to SharePoint Admin Center..." -ForegroundColor Green
Write-Host "  URL: $AdminUrl" -ForegroundColor Green

try {
    if (Test-Path $CertPath) {
        Write-Host "  Using certificate authentication (no browser needed)" -ForegroundColor Green
        Connect-PnPOnline `
            -Url $AdminUrl `
            -ClientId $ClientId `
            -Tenant "$TenantName.onmicrosoft.com" `
            -CertificatePath $CertPath `
            -ErrorAction Stop
    }
    else {
        Write-Host "  No certificate — opening browser for $AdminEmail" -ForegroundColor Green
        Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId -ErrorAction Stop
    }
    Write-Host "✓ Connected!" -ForegroundColor Green
}
catch {
    Write-Host "✗ Connection failed: $($_.Exception.Message)" -ForegroundColor Green
    Write-Host ""
    Write-Host "  • Verify the Client ID is correct for app '$AppName'" -ForegroundColor Green
    Write-Host "  • Ensure the certificate is uploaded to the app in Azure Portal" -ForegroundColor Green
    Write-Host "  • Admin consent must be granted for Sites.FullControl.All" -ForegroundColor Green
    exit
}

# ========== RETRIEVE SharePoint SITES ==========
Write-Host "[DISCOVERY] Scanning for OneDrive sites..." -ForegroundColor Green
Write-Host "(This may take a few minutes depending on tenant size)" -ForegroundColor Green

try {
    $AllSites = Get-PnPTenantSite -IncludeOneDriveSites -ErrorAction Stop
}
catch {
    Write-Host "✗ Failed to retrieve tenant sites!" -ForegroundColor Green
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Possible causes:" -ForegroundColor Green
    Write-Host "  1. The app '$AppName' is missing the Sites.FullControl.All Application permission" -ForegroundColor Green
    Write-Host "     → Azure Portal → App registrations → $AppName → API permissions" -ForegroundColor Green
    Write-Host "     → Add SharePoint → Application → Sites.FullControl.All → Grant admin consent" -ForegroundColor Green
    Write-Host "  2. The certificate is not uploaded to the app" -ForegroundColor Green
    Write-Host "     → Azure Portal → App registrations → $AppName → Certificates & secrets" -ForegroundColor Green
    Write-Host "     → Upload SharePoint-Cleanup-Tool.cer" -ForegroundColor Green
    exit
}

$SharePointSites = $AllSites | Where-Object { $_.Url -like "$MySiteHost/personal/*" }
$Count = $SharePointSites.Count
Write-Host "✓ Found $Count OneDrive site(s)" -ForegroundColor Green

if ($Count -eq 0) {
    Write-Host "✗ No OneDrive sites found." -ForegroundColor Green
    Write-Host "  Retrieved $($AllSites.Count) total site(s) but none matched: $MySiteHost/personal/*" -ForegroundColor Green
    exit
}
Write-Host ""

# ========== USER SELECTION ==========
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "   TARGET USERS — OneDrive owners found in tenant ($Count)  " -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
$i = 0
foreach ($Site in $SharePointSites) {
    $i++
    Write-Host "  [$i] $($Site.Owner)" -ForegroundColor Green
}
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""

$TargetSites = @()

if ($ScopeChoice -eq "S") {
    Write-Host "Select target user(s) by number (whose versions will be deleted):" -ForegroundColor Green
    Write-Host "(Separate multiple numbers with commas, e.g., '1,3,5' or single number '2')" -ForegroundColor Green
    $SelectedNumbers = Read-Host "User number(s)"
    
    # Parse the input numbers
    $NumberArray = $SelectedNumbers -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ }
    
    if ($NumberArray.Count -eq 0) {
        Write-Host "✗ No valid numbers entered. Exiting." -ForegroundColor Green
        exit
    }
    
    # Validate and collect selected sites
    $SelectedEmails = @()
    foreach ($num in $NumberArray) {
        if ($num -ge 1 -and $num -le $SharePointSites.Count) {
            $SelectedEmails += $SharePointSites[$num - 1].Owner
        }
        else {
            Write-Host "⚠ Warning: User number $num is out of range (1-$($SharePointSites.Count)) - skipping" -ForegroundColor Green
        }
    }
    
    if ($SelectedEmails.Count -eq 0) {
        Write-Host "✗ No valid users selected. Exiting." -ForegroundColor Green
        exit
    }
    
    $TargetSites = $SharePointSites | Where-Object { $SelectedEmails -contains $_.Owner }
    
    Write-Host ""
    Write-Host "Selected user(s):" -ForegroundColor Green
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
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "                    ⚠ WARNING ⚠                          " -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "About to DELETE ALL VERSION HISTORY for:" -ForegroundColor Green
Write-Host "  • Users: $($TargetSites.Count)" -ForegroundColor Green
Write-Host "  • This action CANNOT be undone!" -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""
$Confirm = Read-Host "Type 'DELETE' to confirm (or anything else to cancel)"

if ($Confirm -ne "DELETE") {
    Write-Host "✗ Operation cancelled by user." -ForegroundColor Green
    exit
}
Write-Host ""

# ========== PROCESSING LOOP ==========
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "              STARTING VERSION CLEANUP                    " -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "   Parallel processing: $MaxThreads threads" -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""

function Process-UserSite {
    param(
        [string]$Url,
        [string]$Owner,
        [string]$AdminEmail,
        [string]$ClientId,
        [string]$CertPath,
        [string]$TenantName,
        [string]$AdminUrl,
        [int]$ThreadId,
        [hashtable]$SharedStatus = $null
    )
    # Inline status helper — no closure needed
    function Set-Status ([string]$msg, [int]$pct = 0, [bool]$done = $false, [int]$errors = 0) {
        if ($null -ne $SharedStatus) {
            $SharedStatus[$Owner] = @{ msg = $msg; pct = $pct; done = $done; errors = $errors }
        }
    }
    
    $result = [PSCustomObject]@{
        ThreadId        = $ThreadId
        Owner           = $Owner
        Url             = $Url
        Success         = $false
        FilesProcessed  = 0
        VersionsDeleted = 0
        FilesSkipped    = 0
        Errors          = 0
        ErrorMessage    = ""
    }
    
    try {
        $UserConn = $null
        $AdminConn = $null
        Set-Status "Granting site access..." -pct 0
        
        try {
            # Establish a dedicated admin connection inside this runspace
            if (Test-Path $CertPath) {
                $AdminConn = Connect-PnPOnline -Url $AdminUrl -ClientId $ClientId `
                    -Tenant "$TenantName.onmicrosoft.com" -CertificatePath $CertPath `
                    -ReturnConnection -ErrorAction Stop
            }
            Set-PnPTenantSite -Url $Url -Owners @($AdminEmail) -Connection $AdminConn -ErrorAction SilentlyContinue | Out-Null
            Start-Sleep -Seconds 2
        }
        catch { }
        Set-Status "Connecting..." -pct 0
        
        if ((Test-Path $CertPath)) {
            $UserConn = Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant "$TenantName.onmicrosoft.com" -CertificatePath $CertPath -ReturnConnection -ErrorAction Stop
        }
        else {
            # Interactive fallback — only safe for single-threaded use
            $UserConn = Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId -ReturnConnection -ErrorAction Stop
        }
        
        Set-Status "Finding document library..." -pct 0
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
        
        Set-Status "Scanning files..." -pct 0
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
        $fileNum = 0
        
        foreach ($item in $AllFiles) {
            $fileNum++
            $fileName = $item.FieldValues.FileLeafRef
            $fileUrl = $item.FieldValues.FileRef
            
            try {
                $pct = [int](($fileNum / $TotalFileCount) * 100)
                Set-Status "Cleaning file $fileNum / $TotalFileCount - $fileName" -pct $pct
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
        
        Set-Status "Done | $TotalFileCount files | $TotalVersionsCount versions deleted | $ErrorCount errors" -pct 100 -done $true -errors $ErrorCount
        $result.Success = $true
        $result.FilesProcessed = $TotalFileCount
        $result.VersionsDeleted = $TotalVersionsCount
        $result.FilesSkipped = $UserFilesSkipped
        $result.Errors = $ErrorCount
    }
    catch {
        $result.ErrorMessage = $_.Exception.Message
        Set-Status "FAILED: $($_.Exception.Message.Substring(0, [Math]::Min(50, $_.Exception.Message.Length)))" -pct 0 -done $true -errors 1
    }
    
    return $result
}

$TotalFilesProcessed = 0
$TotalVersionsDeleted = 0
$FailedSites = @()
$CompletedSites = @()
$UserNumber = 0

# Capture function definition so it can be injected into each runspace (runspaces don't inherit parent scope)
$ProcessUserSiteFuncDef = ${function:Process-UserSite}.ToString()

$runspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
$runspacePool.Open()

$jobs = [System.Collections.Generic.List[object]]::new()

# Shared live-status table (threads write into it; main loop reads it for display)
$statusTable = [hashtable]::Synchronized(@{})
foreach ($Site in $TargetSites) { $statusTable[$Site.Owner] = @{ msg = "Queued"; pct = 0 } }

Write-Host "[PROCESSING] Queuing $($TargetSites.Count) user(s) across $MaxThreads threads..." -ForegroundColor Green
Write-Host ""

foreach ($Site in $TargetSites) {
    $UserNumber++
    $Url = $Site.Url
    $Owner = $Site.Owner

    Write-Host "  + $Owner" -ForegroundColor Green

    $powershell = [powershell]::Create().AddScript({
            param($Url, $Owner, $AdminEmail, $ClientId, $CertPath, $TenantName, $AdminUrl, $ThreadId, $FuncDef, $SharedStatus)
            # Ensure PnP module is available in this isolated runspace
            Import-Module PnP.PowerShell -ErrorAction SilentlyContinue
            Set-Item -Path "Function:\Process-UserSite" -Value $FuncDef
            Process-UserSite -Url $Url -Owner $Owner -AdminEmail $AdminEmail -ClientId $ClientId -CertPath $CertPath -TenantName $TenantName -AdminUrl $AdminUrl -ThreadId $ThreadId -SharedStatus $SharedStatus
        }).AddParameter("Url", $Url).AddParameter("Owner", $Owner).AddParameter("AdminEmail", $AdminEmail).AddParameter("ClientId", $ClientId).AddParameter("CertPath", $CertPath).AddParameter("TenantName", $TenantName).AddParameter("AdminUrl", $AdminUrl).AddParameter("ThreadId", $UserNumber).AddParameter("FuncDef", $ProcessUserSiteFuncDef).AddParameter("SharedStatus", $statusTable)

    $powershell.RunspacePool = $runspacePool
    $jobs.Add([PSCustomObject]@{
            PowerShell = $powershell
            Handle     = $powershell.BeginInvoke()
            Owner      = $Owner
            Url        = $Url
        })
}

Write-Host ""
Write-Host "[PROCESSING] Running — live status below" -ForegroundColor Green
Write-Host ""


$jobResults = @()
$completed = 0
$spinChars = @('|', '/', '-', '\')
$spinIdx = 0
$totalUsers = $TargetSites.Count
$E = [char]27

# Calculate max owner name length for alignment
$maxOwnerLen = ($TargetSites | ForEach-Object { $_.Owner.Length } | Measure-Object -Maximum).Maximum
$statusWidth = 90 - $maxOwnerLen - 4  # 4 = space + [ + ] + space
if ($statusWidth -lt 20) { $statusWidth = 20 }
# Bar width matches line 1 total: 1(space)+owner+1(space)+1([)+statusWidth+1(]) = owner+statusWidth+4
# Bar line: 1(space) + bar + 1(space) + 4("100%") = bar needs owner+statusWidth+4 -1 -5 = owner+statusWidth-2
$barWidth = $maxOwnerLen + $statusWidth - 2
if ($barWidth -lt 20) { $barWidth = 20 }

# Reserve lines: 1 overall + 2 per user (name + bar)
$displayLines = 1 + ($totalUsers * 2)
for ($i = 0; $i -lt $displayLines; $i++) { Write-Host "" }

# Save the starting row (cursor is now below reserved area)
$startRow = [Console]::CursorTop - $displayLines

while ($jobs.Count -gt 0) {

    $row = $startRow
    $spin = $spinChars[$spinIdx % 4]; $spinIdx++
    $pct = if ($totalUsers -gt 0) { [int](($completed / $totalUsers) * 100) } else { 100 }

    # ── Overall status line ──────────────────────────────────────
    [Console]::SetCursorPosition(0, $row)
    $overallText = " $spin  $completed / $totalUsers complete  |  Files: $TotalFilesProcessed  |  Versions deleted: $TotalVersionsDeleted"
    Write-Host "$E[2K$E[32m$overallText$E[0m" -NoNewline
    $row++

    # ── Per-user: 2 lines each ──────────────────────────────────
    foreach ($site in $TargetSites) {
        $stObj = if ($statusTable.ContainsKey($site.Owner)) { $statusTable[$site.Owner] } else { @{ msg = "Queued"; pct = 0; done = $false; errors = 0 } }
        if ($stObj -is [string]) { $stMsg = $stObj; $stPct = 0; $isDone = $false; $errCount = 0 }
        else { $stMsg = $stObj.msg; $stPct = $stObj.pct; $isDone = $stObj.done; $errCount = $stObj.errors }
        if ($stMsg -like '*Done*') { $stPct = 100 }

        # Line 1: blue username + green [status] — aligned ] + checkmark/X
        [Console]::SetCursorPosition(0, $row)
        $padOwner = $site.Owner.PadRight($maxOwnerLen)
        $padMsg = $stMsg.PadRight($statusWidth).Substring(0, $statusWidth)
        $suffix = ""
        if ($isDone) {
            if ($errCount -eq 0) { $suffix = " $E[32mV$E[0m" }
            else { $suffix = " $E[31mX$E[0m" }
        }
        Write-Host "$E[2K $E[34m$padOwner$E[0m $E[32m[$padMsg]$E[0m$suffix" -NoNewline
        $row++

        # Line 2: per-character gradient bar (red on left -> green on right)
        [Console]::SetCursorPosition(0, $row)
        $uFilled = [int](($stPct / 100) * $barWidth)
        if ($uFilled -gt $barWidth) { $uFilled = $barWidth }
        if ($uFilled -lt 0) { $uFilled = 0 }
        $barStr = ""
        for ($ci = 0; $ci -lt $barWidth; $ci++) {
            $cR = [int](255 - ($ci / [Math]::Max($barWidth - 1, 1)) * 255)
            $cG = [int](($ci / [Math]::Max($barWidth - 1, 1)) * 255)
            $ch = if ($ci -lt $uFilled) { [char]0x2591 } else { '_' }
            $barStr += "$E[38;2;${cR};${cG};0m$ch"
        }
        $pctR = [int](255 - ($stPct * 2.55))
        $pctG = [int]($stPct * 2.55)
        $pctStr = "$stPct%".PadLeft(4)
        Write-Host "$E[2K $barStr$E[0m $E[38;2;${pctR};${pctG};0m$pctStr$E[0m" -NoNewline
        $row++
    }

    # ── Collect completed jobs ──────────────────────────────────────
    # Move cursor below the display area before any Write-Host output from jobs
    [Console]::SetCursorPosition(0, $startRow + $displayLines)

    foreach ($job in @($jobs)) {
        if ($job.Handle.IsCompleted) {
            try {
                $result = $job.PowerShell.EndInvoke($job.Handle)
                $jobResults += $result
                $completed++

                if ($result.Success) {
                    $TotalFilesProcessed += $result.FilesProcessed
                    $TotalVersionsDeleted += $result.VersionsDeleted
                    $CompletedSites += [PSCustomObject]@{
                        Owner           = $result.Owner
                        Url             = $result.Url
                        FilesProcessed  = $result.FilesProcessed
                        VersionsDeleted = $result.VersionsDeleted
                        FilesSkipped    = $result.FilesSkipped
                        Errors          = $result.Errors
                    }
                }
                else {
                    $FailedSites += [PSCustomObject]@{
                        Owner = $result.Owner
                        Url   = $result.Url
                        Error = $result.ErrorMessage
                    }
                }
            }
            catch {
                $completed++
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

    if ($jobs.Count -gt 0) { Start-Sleep -Milliseconds 500 }
}

# Move below display area
[Console]::SetCursorPosition(0, $startRow + $displayLines)
Write-Host ""

$runspacePool.Close()
$runspacePool.Dispose()

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "                   PROCESSING COMPLETE                      " -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host ""

if ($CompletedSites.Count -gt 0) {
    Write-Host "┌─────────────────────────────────────────────────────────┐" -ForegroundColor Green
    Write-Host "│                  SUCCESSFUL SITES                       │" -ForegroundColor Green
    Write-Host "└─────────────────────────────────────────────────────────┘" -ForegroundColor Green
    
    foreach ($site in $CompletedSites) {
        Write-Host "  User:      $($site.Owner)" -ForegroundColor Green
        Write-Host "  Files:     $($site.FilesProcessed)" -ForegroundColor Green
        Write-Host "  Versions:  $($site.VersionsDeleted)" -ForegroundColor Green
        Write-Host "  Skipped:   $($site.FilesSkipped)" -ForegroundColor Green
        Write-Host "  Errors:    $($site.Errors)" -ForegroundColor Green
        Write-Host "  ----------------------------------------" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "  TOTAL FILES PROCESSED:    $TotalFilesProcessed" -ForegroundColor Green
    Write-Host "  TOTAL VERSIONS DELETED:   $TotalVersionsDeleted" -ForegroundColor Green
    Write-Host ""
}

# ========== FINAL SUMMARY ==========
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "                   ✓ JOB COMPLETE ✓                      " -ForegroundColor Green
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "Users Targeted:      $($TargetSites.Count)" -ForegroundColor Green
Write-Host "Files Processed:     $TotalFilesProcessed" -ForegroundColor Green
Write-Host "Versions Deleted:    $TotalVersionsDeleted" -ForegroundColor Green
Write-Host "Failed Sites:        $($FailedSites.Count)" -ForegroundColor $(if ($FailedSites.Count -gt 0) { "Red" } else { "Green" })
Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Green

if ($FailedSites.Count -gt 0) {
    Write-Host ""
    Write-Host "┌──────────────────────────────────────────────────────┐" -ForegroundColor Green
    Write-Host "│                  FAILED SITES                        │" -ForegroundColor Green
    Write-Host "└──────────────────────────────────────────────────────┘" -ForegroundColor Green
    foreach ($Failed in $FailedSites) {
        Write-Host "  User:  $($Failed.Owner)" -ForegroundColor Green
        Write-Host "  Error: $($Failed.Error)" -ForegroundColor Green
        Write-Host "  ───────────────────────────────────────" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "Script completed at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Green



