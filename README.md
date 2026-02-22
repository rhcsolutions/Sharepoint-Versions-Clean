# SharePoint Version Cleanup Tool

Automated PowerShell script to bulk-delete file version history from OneDrive personal sites across a Microsoft 365 tenant.

## Overview

- **Parallel processing** — processes multiple users simultaneously (default: 10 threads)
- **Certificate authentication** — no browser popup, connects silently using the included `.pfx`
- **Remembers your settings** — admin account, tenant, and app ID saved to `config.json` after first run
- **Repeat runs: one keypress** — press Enter to reuse saved settings, or `N` to switch tenant/app
- **Smart app registration** — creates a new Azure AD app with all required permissions, or reuses existing
- **Auto-verifies permissions** — patches `Sites.FullControl.All` via Graph API every connection
- **Bulk version deletion** — scans all files per user, deletes all versions with retry logic
- **Verbose output** — shows progress for every file being processed
- **Per-user + final summary** reported at the end

## Requirements

| Requirement | Details |
| --- | --- |
| PowerShell | 5.1 or higher |
| PnP PowerShell | `Install-Module PnP.PowerShell` (v3.x recommended) |
| Account role | Global Administrator or SharePoint Administrator |
| Certificate | `SharePoint-Cleanup-Tool.pfx` in the script folder |
| App permissions | `AllSites.FullControl` (Delegated) + `Sites.FullControl.All` (Application) |

## Quick Start

### 1. Install PnP PowerShell

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

### 2. Run the script

```powershell
.\CleanSharepointVersions.ps1
```

To use a different number of threads (default is 10):

```powershell
.\CleanSharepointVersions.ps1 -MaxThreads 20
```

### 3. First run — answer 2 questions

```text
[STEP 1] Admin account:   admin@contoso.onmicrosoft.com
         (tenant extracted automatically)

[STEP 2] App registration:
         [Enter] Use existing app   [N] Create new app

[STEP 3] Target users:
         [A] ALL  or  [S] SPECIFIC
```

### 4. Repeat runs — one keypress

```text
  Saved settings found:
    Admin:    admin@contoso.onmicrosoft.com
    Tenant:   contoso.onmicrosoft.com
    App ID:   xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx

  [Enter] Use saved settings
  [N]     Enter new settings (different tenant / new app)
  Choice:
```

Press **Enter** — skips directly to user selection.

## Files

| File | Purpose |
| --- | --- |
| `CleanSharepointVersions.ps1` | Main script |
| `fix_encoding.ps1` | Utility to fix encoding issues (run if script fails to parse) |
| `SharePoint-Cleanup-Tool.pfx` | Certificate for app authentication |
| `SharePoint-Cleanup-Tool.cer` | Public cert — upload to Azure AD app |
| `config.json` | Saved settings (gitignored, created on first run) |

> `config.json` is excluded from git via `.gitignore` — it stays local only.

## App Registration

| Scenario | What the script does |
| --- | --- |
| No saved Client ID | Prompts to create a new app, registers with both permissions + cert, opens consent page |
| Saved Client ID exists | Shows saved ID — press Enter to reuse, or `N` to create a new app |
| After every connection | PATCHes the app via Graph API to ensure both permissions are set |
| Graph PATCH fails | Shows manual Azure Portal steps with exact navigation path |

### Required Permissions

| Type | Permission | Used for |
| --- | --- | --- |
| Delegated | `AllSites.FullControl` | Access user OneDrive sites |
| Application | `Sites.FullControl.All` | `Get-PnPTenantSite` tenant-wide discovery |

> Both permissions require **admin consent**. The script opens the consent URL automatically when creating a new app.

## Certificate Setup

The `.pfx` and `.cer` are already included in this repo folder.

To upload the certificate to an **existing** app:

1. Azure Portal → App registrations → `SharePoint-Cleanup-Tool`
2. Certificates & secrets → Upload certificate
3. Select `SharePoint-Cleanup-Tool.cer` → Add

When **creating a new app** via `[N]`, the script registers it with the `.pfx` automatically.

## Processing Workflow

```text
[Startup]
  Load config.json → show saved settings → [Enter] reuse / [N] re-enter
       ↓
[STEP 1] Admin email → tenant name extracted automatically
       ↓
[STEP 2] Certificate found → connect using cert (no browser)
         App ID confirmed or new app created → admin consent
       ↓
[STEP 3] Choose ALL or SPECIFIC users
       ↓
[CONNECTION] Connect-PnPOnline with certificate
[PERMISSIONS] Graph API PATCH → ensure Sites.FullControl.All
[DISCOVERY] Get-PnPTenantSite → list all OneDrive personal sites
       ↓
If [S]: pick user(s) from numbered green list
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

## Output Example

```text
  Saved settings found:
    Admin:    admin@contoso.onmicrosoft.com
    Tenant:   contoso.onmicrosoft.com
    App ID:   ca3271dd-431a-4bb9-b4f2-4583577d4def
  Choice: [Enter]

  ✓ Certificate found: SharePoint-Cleanup-Tool.pfx
  ✓ App ID: ca3271dd-431a-4bb9-b4f2-4583577d4def

[STEP 3] Target Users:  [A] ALL  or  [S] SPECIFIC
Choice (A/S): s

[CONNECTION] Connecting to SharePoint Admin Center...
  Using certificate authentication (no browser needed)
✓ Connected!

[PERMISSIONS] Verifying app permissions...
✓ Permissions confirmed: AllSites.FullControl (Delegated) + Sites.FullControl.All (Application)

[DISCOVERY] Scanning for OneDrive sites...
✓ Found 12 OneDrive site(s)

══════════════════════════════════════════════════════════
   TARGET USERS — OneDrive owners found in tenant (12)
══════════════════════════════════════════════════════════
  [1]  user1@contoso.onmicrosoft.com
  [2]  user2@contoso.onmicrosoft.com
  ...

┌─────────────────────────────────────────────────────────┐
│ USER [1/2]: user1@contoso.onmicrosoft.com               │
└─────────────────────────────────────────────────────────┘
  [1/6] Granting admin access... ✓
  [2/6] Connecting to user site... ✓
  [3/6] Detecting document library... ✓ 'Documents'
  [4/6] Setting version limit to 1... ✓
  [5/6] Scanning files... ✓ Found 1,245 file(s)
  [6/6] Deleting versions...

  Processing 1,245 file(s) — verbose output below:
  [1/1245] report.xlsx  → ✓ 3 version(s) deleted
  [2/1245] budget.xlsx  → ✓ 5 version(s) deleted
  [3/1245] notes.docx  → skipped (no versions)
  ...

  ┌───────────────────────────────────────────┐
  │          USER PROCESSING SUMMARY          │
  │ Total Files:                         1245 │
  │ Cleaned:                             1240 │
  │ Skipped:                                5 │
  │ Errors:                                 0 │
  └───────────────────────────────────────────┘
```

## Troubleshooting

### Script fails to parse with "Missing closing" errors

The PowerShell script file may have encoding issues. Run the included utility:

```powershell
.\fix_encoding.ps1
```

Then re-run the main script.

### "Attempted to perform an unauthorized operation" on `Get-PnPTenantSite`

The app is missing `Sites.FullControl.All` Application permission, or the certificate is not uploaded.

The script auto-fixes this via Graph API after every connection. If it still fails:

1. Azure Portal → App registrations → `SharePoint-Cleanup-Tool` → API permissions
2. Add a permission → SharePoint → Application permissions → `Sites.FullControl.All`
3. Grant admin consent for your tenant

### "Connection failed" — certificate error

- Verify `SharePoint-Cleanup-Tool.cer` is uploaded to the app in Azure Portal
- Confirm the Client ID in `config.json` matches the app in Azure Portal
- Ensure admin consent has been granted

### "No OneDrive sites found" after successful connection

- Admin consent for `Sites.FullControl.All` may not be fully propagated — wait 2–5 minutes and retry
- Verify the account has **Global Administrator** or **SharePoint Administrator** role

### Want to use a different tenant?

Press **`N`** at the saved settings prompt — re-enters STEP 1 and STEP 2 with fresh values, saved to `config.json` at the end.

## Important Notes

- **Version deletion is permanent** — there is no undo
- Type **`DELETE`** exactly at the confirmation prompt to proceed
- Admin access is **kept** on processed sites by default (`$RemoveAccessAfter = $false`)
- Logging is **VERBOSE** by default — every file and version count is shown
- Large libraries (100k+ files) automatically fall back to a faster scan method

---

## © RHC Solutions

Developed and maintained by **RHC Solutions**.

| | |
| --- | --- |
| Website | [rhcsolutions.com](https://rhcsolutions.com) |
| Telegram | [t.me/rhcsolutions](https://t.me/rhcsolutions) |

© 2026 RHC Solutions. All rights reserved.
