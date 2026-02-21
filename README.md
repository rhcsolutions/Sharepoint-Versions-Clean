# SharePoint Version Cleanup Tool

An automated PowerShell script to clean up version history from SharePoint Online and OneDrive files at scale.

## 📋 Overview

This tool helps SharePoint administrators efficiently delete all version history from user files across the organization. It:

- **Automatically manages Azure AD app registration** (creates new or detects existing)
- **Discovers all OneDrive personal sites** in your tenant
- **Allows selective or bulk user processing** via number-based selection
- **Deletes file versions in bulk** (optimized for performance)
- **Provides real-time progress tracking** with file names
- **Reports comprehensive statistics** and failures
- **Optionally removes admin access** after completion

## ⚙️ Requirements

- **PowerShell 5.1 or higher**
- **PnP PowerShell module** (`Install-Module PnP.PowerShell`)
- **SharePoint Admin role** in Microsoft 365 tenant
- **Browser access** (for interactive authentication and Azure Portal)
- **Admin consent rights** for Azure AD app registration

## 🚀 Quick Start

### 1. Install PnP PowerShell Module

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

### 2. Run the Script

```powershell
# Open PowerShell and run:
.\CleanSharepointVersions
```

### 3. Follow the Interactive Prompts

The script will guide you through:
- **STEP 1:** Enter tenant name (e.g., "contoso" from contoso.onmicrosoft.com)
- **STEP 2:** Azure AD app registration (automatic creation or detection)
- **STEP 3:** Administrator email address
- **STEP 4:** Select processing scope (ALL users or SPECIFIC users)
- **STEP 5:** Choose logging verbosity level
- **STEP 6:** Decide whether to remove admin access after cleanup
- **FINAL:** Review configuration and confirm with "DELETE" to proceed

## 📊 Features

### Azure AD App Registration
- **Automatic Detection:** Checks if "SharePoint-Cleanup-Tool" app already exists
- **Smart Creation:** Creates app with required "AllSites.FullControl" permission if needed
- **Fallback to Portal:** Opens Azure Portal if automatic retrieval fails
- **Manual Entry:** Allows copying Client ID directly from Azure Portal

### User Selection
- **Number-Based Selection:** Select users by number instead of email addresses
  - Example: Enter "1,3,5" to process users 1, 3, and 5
  - Or enter "2" for a single user
- **Bulk Option:** Process all users at once with one selection
- **Confirmation Display:** Shows selected users before processing

### Performance Optimization
- **Bulk Deletion:** Deletes all file versions at once (like the web interface)
- **Smart Progress:** Shows file names every 10 files instead of every file
- **Batch Processing:** Handles large libraries with timeout protection
- **Retry Logic:** Automatically retries transient failures up to 3 times

### Error Handling
- **Graceful Fallbacks:** Multiple methods to detect document libraries (multi-language support)
- **Timeout Protection:** Falls back to fast scan if library is too large
- **Detailed Logging:** Reports specific errors with file names and details
- **Failed Site Tracking:** Lists all failed sites at the end with error reasons

### Security
- **Admin Access Management:** Optionally removes temporary admin access after cleanup
- **Interactive Authentication:** Uses browser-based login (no hardcoded credentials)
- **Permission Validation:** Confirms required permissions before proceeding

## 📈 Processing Workflow

```
1. Tenant Configuration
2. Azure AD App Registration (Create or Detect)
3. Admin Consent (if new app)
4. Connect to SharePoint Admin Center
5. Discover OneDrive Sites
6. Select Target User(s)
7. For Each Selected User:
   - Grant admin access
   - Connect to OneDrive
   - Detect document library
   - Set version limit to 1
   - Scan all files
   - Delete all versions (bulk operation)
   - Remove admin access (optional)
8. Generate Final Report
```

## 📊 Output Example

```
[USER 1/5]: user1@contoso.com
URL: https://contoso-my.sharepoint.com/personal/user1_contoso_com

  [1/6] Granting admin access... ✓
  [2/6] Connecting to user site... ✓
  [3/6] Detecting document library... ✓ 'Documents'
  [4/6] Setting version limit to 1... ✓
  [5/6] Scanning files... ✓ Found 1,245 file(s)
  [6/6] Deleting versions...

        Processing 1,245 file(s) - Deleting all versions at once...

        [10/1245 - 0.8%] document.docx
        [20/1245 - 1.6%] spreadsheet.xlsx
        [30/1245 - 2.4%] presentation.pptx

  ┌───────────────────────────────┐
  │   USER PROCESSING SUMMARY     │
  ├───────────────────────────────┤
  │ Total Files:              1245 │
  │ Cleaned:                  1240 │
  │ Skipped:                     5 │
  │ Errors:                      0 │
  └───────────────────────────────┘
```

## 🔑 Configuration Summary

The script displays a summary before processing:

```
==================================================
              CONFIGURATION SUMMARY
==================================================
Tenant:        contoso.onmicrosoft.com
Admin:         admin@contoso.com
App ID:        xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
Scope:         ALL USERS (or SPECIFIC USERS)
Logging:       NORMAL (or VERBOSE)
Remove Access: YES (or NO)
==================================================
```

## ⚠️ Important Notes

- **Cannot Be Undone:** Version history deletion is permanent. There is no undo.
- **Admin Confirmation Required:** Type "DELETE" exactly to confirm execution
- **Large Libraries:** Libraries with 100,000+ files may take considerable time
- **Network Stability:** Long-running operations should use a stable connection
- **Admin Consent:** New apps require admin consent in Azure Portal before first use
- **Timeout Protection:** Script handles library timeouts with fallback methods

## 🐛 Troubleshooting

### "Client ID is required" Error
- Ensure Azure AD app has admin consent granted
- Verify Client ID is correct in Azure Portal
- Check that app permissions include "AllSites.FullControl"

### "Connection failed" Error
- Verify your account has **SharePoint Admin** role
- Check that the tenant name is correct
- Ensure you have network access to SharePoint
- Try using InPrivate/Incognito browser window

### "No OneDrive sites found"
- Verify admin credentials have proper permissions
- Check that OneDrive is provisioned for users in your tenant
- Wait a few minutes and try again (new sites may not be immediately discoverable)

### "Library scan timeout"
- Very large libraries (100,000+ files) may timeout
- Consider processing library directly in SharePoint web interface
- Or contact Microsoft Support for large-scale cleanup

### "No file versions to delete"
- Versioning may be disabled on the library
- Files may already have no version history
- Check library settings in SharePoint

## 📝 Log Output

The script displays:
- Progress with file names (every 10 files)
- Per-user statistics (Files, Cleaned, Skipped, Errors)
- Final summary with total counts
- Failed sites with specific error reasons

## 🔐 Security Best Practices

1. **Minimal Permissions:** App uses delegated permissions (on behalf of user)
2. **Admin Access Cleanup:** Script removes temporary admin rights after processing
3. **No Hardcoded Credentials:** Uses interactive browser-based authentication
4. **Audit Trail:** Consider enabling SharePoint audit logging before cleanup

## 📦 Azure AD App Permissions

Required permission: **AllSites.FullControl** (Delegated)

This permission is automatically configured when the script creates the app.

## 🆘 Getting Help

If you encounter issues:

1. **Check tenant name format:** Should be the subdomain (e.g., "contoso" not "contoso.com")
2. **Verify admin role:** Account must have SharePoint Admin role
3. **Test app registration:** Manually verify app in Azure Portal
4. **Review error messages:** Script provides specific error details
5. **Check failed sites report:** Details which specific sites had issues

## 📄 License

[Your License Here]

## 👤 Author

SharePoint Admin Team  
December 2024

---

## Script Flowchart

```
Start
  ↓
[1] Tenant Name Input
  ↓
[2] Azure AD App (Create/Detect)
  ↓
[3] Admin Email Input
  ↓
[4] Processing Scope (ALL/SPECIFIC)
  ↓
[5] Logging Verbosity
  ↓
[6] Remove Access After? (Y/N)
  ↓
Configuration Summary → Confirm with "DELETE"
  ↓
Connect to Admin Center
  ↓
Discover OneDrive Sites
  ↓
Select User(s)
  ↓
For Each User:
  ├─ Grant Admin Access
  ├─ Connect to OneDrive
  ├─ Detect Library
  ├─ Set Version Limit
  ├─ Scan Files
  ├─ Delete All Versions (Bulk)
  └─ Remove Admin Access (optional)
  ↓
Generate Final Report
  ↓
End
```

---

**Last Updated:** December 2025  
**Script Version:** 1.0  
**Status:** Production Ready
