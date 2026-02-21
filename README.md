# SharePoint Version Cleanup Tool

An automated PowerShell script to delete file version history from OneDrive personal sites across a Microsoft 365 tenant.

## 📋 Overview

This tool allows SharePoint administrators to bulk-delete version history from user OneDrive files. It:

- **Asks only 2 questions** — admin account + which users to target
- **Auto-extracts the tenant name** from your admin email
- **Automatically creates or detects** the required Azure AD app registration
- **Auto-updates app permissions** if the app exists but is missing required permissions
- **Discovers all OneDrive personal sites** in the tenant
- **Deletes all file versions in bulk** per user
- **Reports per-user and final statistics**

## ⚙️ Requirements

| Requirement | Details |
| --- | --- |
| PowerShell | 5.1 or higher |
| PnP PowerShell | `Install-Module PnP.PowerShell` (v3.x) |
| Account role | **Global Administrator** or **SharePoint Administrator** |
| App permissions | `AllSites.FullControl` (Delegated) + `Sites.FullControl.All` (Application) |
| Browser | Required for interactive sign-in and admin consent |

## 🚀 Quick Start

### 1. Install PnP PowerShell

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

### 2. Run the Script

```powershell
.\SharePoint-Cleanup-Tool.ps1
```

### 3. Answer 2 Questions

```text
[STEP 1] Admin account:   admin@contoso.onmicrosoft.com
[STEP 2] App registration: auto-detected / created
[STEP 3] Target users:    [A] ALL  or  [S] SPECIFIC
```

The script handles everything else automatically.

## 🔄 What Happens at Each Run

### App Registration (STEP 2)

| Scenario | What the script does |
| --- | --- |
| App **does not exist** | Creates it with correct permissions, opens consent page |
| App **exists** | Retrieves Client ID, updates permissions via Graph API, re-grants consent |
| Auto-update **fails** | Opens Azure Portal API permissions page with step-by-step instructions |

### Required App Permissions

| Type | Permission | Used for |
| --- | --- | --- |
| Delegated | `AllSites.FullControl` | Access user OneDrive sites on your behalf |
| Application | `Sites.FullControl.All` | `Get-PnPTenantSite -IncludeOneDriveSites` (tenant discovery) |

> ⚠️ Both permissions require **admin consent** in Azure Portal.

## 📈 Processing Workflow

```text
STEP 1 → Enter admin email (tenant name auto-extracted)
STEP 2 → App found/created → permissions verified → admin consent
STEP 3 → Choose ALL users or SPECIFIC users
       ↓
Sign in via browser (Interactive auth)
       ↓
Auto-update app permissions via Graph API (if app pre-existed)
Re-grant admin consent
       ↓
Discover all OneDrive personal sites
If [S]: pick users from numbered green list
       ↓
For each target user:
  [1/6] Grant admin access to their OneDrive
  [2/6] Connect to their OneDrive site
  [3/6] Detect document library (multi-language support)
  [4/6] Set version limit to 1
  [5/6] Scan all files
  [6/6] Delete all versions (bulk, with retry logic)
       ↓
Final summary report
```

## 📊 Output Example

```text
[CONNECTION] Sign in with your admin account in the browser...
  Account: admin@contoso.onmicrosoft.com
✓ Connected!

[PERMISSIONS] Ensuring app has required permissions...
✓ Permissions updated (Delegated: AllSites.FullControl + Application: Sites.FullControl.All)

[DISCOVERY] Scanning for OneDrive sites...
✓ Found 12 OneDrive site(s)

══════════════════════════════════════════════════════════
   TARGET USERS — OneDrive owners found in tenant (12)
══════════════════════════════════════════════════════════
  [1] user1@contoso.onmicrosoft.com
  [2] user2@contoso.onmicrosoft.com
  ...

┌─────────────────────────────────────────────────────────┐
│ USER [1/3]: user1@contoso.onmicrosoft.com               │
└─────────────────────────────────────────────────────────┘
  [1/6] Granting admin access... ✓
  [2/6] Connecting to user site... ✓
  [3/6] Detecting document library... ✓ 'Documents'
  [4/6] Setting version limit to 1... ✓
  [5/6] Scanning files... ✓ Found 1,245 file(s)
  [6/6] Deleting versions...

  ┌───────────────────────────────────────────┐
  │          USER PROCESSING SUMMARY          │
  │ Total Files:                         1245 │
  │ Cleaned:                             1240 │
  │ Skipped:                                5 │
  │ Errors:                                 0 │
  └───────────────────────────────────────────┘
```

## 🐛 Troubleshooting

### "Attempted to perform an unauthorized operation" on `Get-PnPTenantSite`

The app is missing the **Application** permission `Sites.FullControl.All`.

Re-run the script — it will auto-fix this after you sign in.

Or add it manually: Azure Portal → App Registrations → `SharePoint-Cleanup-Tool` → API permissions → Add `SharePoint > Application > Sites.FullControl.All` → Grant admin consent.

### "Connection failed" / browser doesn't open

- Verify the account has **Global Administrator** or **SharePoint Administrator** role
- Make sure the Client ID is correct
- Try running PowerShell as Administrator

### "No OneDrive sites found" after successful connection

- Admin consent may not have been granted for `Sites.FullControl.All`
- Re-run the script — it will re-grant consent automatically
- Wait 2–5 minutes after granting consent before retrying

### App auto-permission update fails

The script opens the Azure Portal API Permissions page and shows step-by-step instructions:

1. Add a permission → SharePoint → Application permissions → `Sites.FullControl.All`
2. Grant admin consent for your tenant

## ⚠️ Important Notes

- **Version deletion is permanent** — there is no undo
- Type **`DELETE`** exactly at the confirmation prompt to proceed
- Admin access is **kept** on processed sites by default (`$RemoveAccessAfter = $false`)
- Logging is **VERBOSE** by default — every file and version count is shown
- Large libraries (100k+ files) use an automatic fallback to a faster scan method
