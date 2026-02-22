# SharePoint Version Cleanup Tool

Bulk-deletes file version history from OneDrive personal sites across a Microsoft 365 tenant. Runs in parallel with a real-time ANSI progress display.

## Features

- **Parallel processing** — multiple users processed simultaneously via PowerShell runspaces (default: 10 threads, configurable with `-MaxThreads`)
- **Live ANSI dashboard** — in-place updating display with per-user status, per-character gradient progress bar (red to green), and overall counters
- **Certificate authentication** — connects silently using a `.pfx` certificate, no browser popups
- **Saved settings** — admin account, tenant, and app ID persisted to `config.json` for one-keypress repeat runs
- **Smart app registration** — creates Azure AD app with required permissions and certificate, or reuses existing
- **Multi-language library detection** — finds the document library regardless of locale (Documents, Documenten, Dokumenty, Documentos, etc.)
- **Retry logic** — failed version deletions are retried up to 3 times before marking as error
- **Completion indicators** — green V on success, red X on errors per user

## Requirements

| Requirement | Details |
| --- | --- |
| PowerShell | 5.1 or higher |
| PnP PowerShell | `Install-Module PnP.PowerShell` (v3.x recommended) |
| Account role | SharePoint Administrator |
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

Custom thread count:

```powershell
.\CleanSharepointVersions.ps1 -MaxThreads 20
```

### 3. First run

The script walks through 3 steps:

1. **Admin account** — enter your SharePoint admin email (tenant name extracted automatically)
2. **App registration** — create a new Azure AD app or enter an existing Client ID
3. **Target users** — choose ALL users or pick specific ones from a numbered list

Settings are saved to `config.json` for next time.

### 4. Repeat runs

```
  Saved settings found:
    Admin:    admin@contoso.onmicrosoft.com
    Tenant:   contoso.onmicrosoft.com
    App ID:   xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx

  [Enter] Use saved settings
  [N]     Enter new settings
```

Press **Enter** to skip straight to user selection.

## Live Display

During processing, the script shows a real-time dashboard that updates in-place:

```
 /  3 / 10 complete  |  Files: 1245  |  Versions deleted: 3891
 user1@contoso.com  [42/500 /Documents/Reports/Q4/budget.xlsx                          ]
 ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░__________________________________                8%
 user2@contoso.com  [Done | 800 files | 2100 versions deleted | 0 errors               ] V
 ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░ 100%
 user3@contoso.com  [Connecting...                                                      ]
 ___________________________________________________________________________________   0%
```

- **Usernames** in cyan, **status** in green with aligned brackets
- **Progress bar** uses per-character RGB gradient from red (left) to green (right)
- `░` = filled (progress), `_` = unfilled (remaining)
- **V** = completed without errors, **X** = completed with errors
- File paths shown from library root (`/Documents/...`) without the `/personal/username/` prefix
- 200-character wide display

## Files

| File | Purpose |
| --- | --- |
| `CleanSharepointVersions.ps1` | Main script |
| `fix_encoding.ps1` | Utility to fix encoding issues if script fails to parse |
| `SharePoint-Cleanup-Tool.pfx` | Certificate for app authentication |
| `SharePoint-Cleanup-Tool.cer` | Public cert — upload to Azure AD app |
| `config.json` | Saved settings (created on first run, gitignored) |

## App Registration

### Automatic

When no Client ID is saved, the script offers to create a new app via `Register-PnPAzureADApp` with:

- `AllSites.FullControl` (Delegated)
- `Sites.FullControl.All` (Application)
- Certificate automatically attached

It then opens the admin consent URL in your browser.

### Manual

If automatic registration fails, enter an existing Client ID. Ensure the app has:

1. Azure Portal > App registrations > `SharePoint-Cleanup-Tool` > API permissions
2. SharePoint > `AllSites.FullControl` (Delegated) + `Sites.FullControl.All` (Application)
3. Admin consent granted
4. Certificate uploaded under Certificates & secrets

## Processing Workflow

```
Startup > Load config.json > [Enter] reuse / [N] re-enter
   |
STEP 1: Admin email > tenant extracted
   |
STEP 2: Certificate check > App ID confirmed or created > admin consent
   |
STEP 3: Choose ALL or SPECIFIC users
   |
Connect to SharePoint Admin Center (certificate auth)
   |
Discovery: Get-PnPTenantSite > list OneDrive personal sites
   |
User selection (if SPECIFIC): pick from numbered list
   |
Confirmation: type DELETE to proceed
   |
Parallel processing (runspace pool):
  Per user:
    - Grant admin access to OneDrive site
    - Connect to user site
    - Detect document library (multi-language)
    - Set version limit to 1
    - Scan all files
    - Delete all versions (retry up to 3x per file)
   |
Final summary: files processed, versions deleted, failures
```

## Troubleshooting

### Script fails to parse

Run the encoding fix utility, then re-run:

```powershell
.\fix_encoding.ps1
```

### "Attempted to perform an unauthorized operation" on Get-PnPTenantSite

The app is missing `Sites.FullControl.All` Application permission:

1. Azure Portal > App registrations > `SharePoint-Cleanup-Tool` > API permissions
2. Add permission > SharePoint > Application > `Sites.FullControl.All`
3. Grant admin consent

### Connection failed — certificate error

- Verify `SharePoint-Cleanup-Tool.cer` is uploaded to the app in Azure Portal
- Confirm the Client ID in `config.json` matches the app
- Ensure admin consent has been granted

### No OneDrive sites found

- Admin consent may not be fully propagated — wait 2-5 minutes and retry
- Verify the account has the **SharePoint Administrator** role

### Switch to a different tenant

Press **N** at the saved settings prompt to re-enter all values.

## Important Notes

- **Version deletion is permanent** — there is no undo
- Type **DELETE** exactly at the confirmation prompt to proceed
- Admin access is kept on processed sites by default
- Large libraries (100k+ files) automatically fall back to smaller page sizes
- All terminal output is green (with cyan usernames) for consistent appearance

---

## RHC Solutions

| | |
| --- | --- |
| Website | [rhcsolutions.com](https://rhcsolutions.com) |
| Telegram | [t.me/rhcsolutions](https://t.me/rhcsolutions) |

2026 RHC Solutions
