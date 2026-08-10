# Microsoft Entra ID Guest User Lifecycle Toolkit

| | |
|---|---|
| **Author** | Waverly Chua |
| **Creation Date** | 5 Aug 2026 |
| **Version** | 1.0 |

Three PowerShell scripts covering the full guest-user lifecycle: review,
clean up, and recover. Each has its own section below — jump straight
to the one you need, or read start to finish for the full picture.

| Script | Purpose | Modifies accounts? |
|---|---|---|
| [`Get-GuestUserDetails.ps1`](#1-get-guestuserdetailsps1) | Full Excel review report, all guests, 6 categories | No (read-only) |
| [`Get-And-Remove-GuestUsers.ps1`](#2-get-and-remove-guestusersps1) | Criteria-based txt report + optional delete | Only with `-Delete` |
| [`Restore-DeletedGuestUsers.ps1`](#3-restore-deletedguestusersps1) | Recover soft-deleted guests by email or Object ID | Yes (restores) |

**Recommended workflow:** review with `Get-GuestUserDetails.ps1` →
decide on thresholds → clean up with `Get-And-Remove-GuestUsers.ps1
-Delete` → if something was removed by mistake, recover within 30 days
with `Restore-DeletedGuestUsers.ps1`.

---

## 1. Get-GuestUserDetails.ps1

### Overview

Connects to Microsoft Graph and generates an Excel report containing
Microsoft Entra ID guest user information. Read-only — makes no changes
to any account.

### Features

- Checks required modules
- Connects to Microsoft Graph
- Retrieves guest users
- Collects sign-in activity
- Creates categorized worksheets
- Exports formatted Excel report

### Requirements

PowerShell 7+, `Microsoft.Graph` module, `ImportExcel` module

### Permissions

| Scope | Purpose |
|---|---|
| `User.Read.All` | Read guest user profiles |
| `AuditLog.Read.All` | Read `SignInActivity` (last sign-in timestamps) |

### Output

`GuestUserReview_YYYY-MM-DD.xlsx`

### Worksheets

All Guest Users, Inactive >180 Days, Accepted Never Signed In, Pending
Invitation, Pending >90 Days, Recent Invite <30 Days

### How to Run

```powershell
.\Get-GuestUserDetails.ps1
```

Sign in when prompted. The report will be generated automatically.

### Troubleshooting

Install missing modules using:
```powershell
Install-Module Microsoft.Graph
Install-Module ImportExcel
```

---

## 2. Get-And-Remove-GuestUsers.ps1

### Objective

To maintain a secure and clean Microsoft Entra ID (Azure AD) environment
by systematically identifying and managing stale B2B guest accounts.
This script provides an automated, safety-first approach to guest
lifecycle management. It audits your tenant to find guests who are
inactive or have lingering, unaccepted invitations, logs these findings
into a detailed plain-text report for review, and — **only when
explicitly instructed** — safely removes the matching accounts.

### How It Works

1. **Connects to Microsoft Graph:** Reuses an existing session if
   you're already signed in.
2. **Retrieves Guest Users:** Fetches every guest in the tenant by
   default, or only the UPNs/emails listed in a text file via
   `-UserListPath`.
3. **Applies Filters:** Narrows down the list to guests matching your
   criteria:
   - **Inactive:** Invite Accepted, but no sign-in in more than
     `-InactiveDays` (default: 180 days).
   - **Stale Invite:** Invite still Pending Acceptance, sent more than
     `-PendingDays` ago (default: 90 days).
   - *(Anyone matching neither condition is excluded entirely — not
     reported, not touched.)*
4. **Generates a Report:** Writes a plain `.txt` file with full details
   for every matching guest, including a `MatchReason`.
5. **Stops or Proceeds:**
   - **Default:** Stops here. No accounts are removed unless you pass
     the `-Delete` flag.
   - **If `-Delete` is passed:** Prompts for a `Y/N` confirmation
     (unless `-Force` is passed), then removes each matching guest one
     at a time. Skips any account that turns out not to actually be a
     Guest.
6. **Logs Removals:** Writes a CSV removal report tracking `Success`,
   `Failed`, or `Skipped` status per guest.

### Prerequisites & Permissions

**System Requirements:**
```powershell
Install-Module Microsoft.Graph.Users -Scope CurrentUser
```
**Licensing:** Azure AD Premium P1 or P2 license — **required** for the
`SignInActivity` data the inactivity check relies on. Without it,
sign-in fields return empty, and every guest may falsely appear
"inactive".

**Graph API Scopes:**

| Scope | Purpose |
|---|---|
| `User.ReadWrite.All` | Read guest profiles, and delete them if `-Delete` is used |
| `AuditLog.Read.All` | Read `SignInActivity` for the inactivity check and text report |

> Note: Microsoft Graph doesn't have a standalone `User.Write` scope —
> the script requests one combined scope, `User.ReadWrite.All`, which
> covers both reading and writing/deleting guest profiles.

**Entra ID Directory Role:** the signed-in account also needs one of:
- **User Administrator**
- **Global Administrator**

*(Without one of these roles, `-Delete` fails with a permission error
even if sign-in and scope consent succeed.)*

### Usage Examples

**Report only (Default) — Always start here.** Writes a text file
listing every guest matching the default criteria (Inactive > 180 days
OR Pending > 90 days). No accounts are touched.
```powershell
.\Get-And-Remove-GuestUsers.ps1
```

**Custom thresholds:**
```powershell
.\Get-And-Remove-GuestUsers.ps1 -InactiveDays 90 -PendingDays 30
```

**Scope to a specific list of guests** (criteria still applies on top
of the list — a listed guest who's actually still active or recently
invited is excluded from the report):
```powershell
.\Get-And-Remove-GuestUsers.ps1 -UserListPath "C:\Temp\guests.txt"
```

**Report + Delete, with confirmation prompt:**
```powershell
.\Get-And-Remove-GuestUsers.ps1 -Delete
```

**Report + Delete, fully unattended (no prompt)** — for scheduled/
automated runs (Azure Automation, Power Automate, Task Scheduler):
```powershell
.\Get-And-Remove-GuestUsers.ps1 -Delete -Force
```

> **Note on Automation:** This script uses interactive `Connect-MgGraph`
> sign-in by default. For a truly unattended scheduled run, switch
> authentication to app-only (certificate or client secret) inside the
> script — otherwise it will hang waiting for a login prompt that never
> comes.

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-UserListPath` | *(none — all guests)* | Path to a `.txt` file of specific UPNs/emails to scope to, one per line. |
| `-DetailsPath` | `guest-details_<timestamp>.txt` | Where the txt detail report is written (saves to script folder). |
| `-ReportPath` | `GuestUserRemoval_Report_<timestamp>.csv` | Where the removal CSV report is written (only created if `-Delete` is used). |
| `-InactiveDays` | `180` | Days since last sign-in before an Accepted guest counts as inactive. |
| `-PendingDays` | `90` | Days after invite before a Pending guest counts as stale. |
| `-Delete` | *(off)* | Remove matching guests. Without this, the script is report-only. |
| `-Force` | *(off)* | Skip the `Y/N` confirmation prompt when `-Delete` is used. |

**Guest UPN Sample:** `johndoe@gmail.com#EXT#@yourtenant.onmicrosoft.com`

### Output Files

**1. Detail Report (`.txt`)** — one block per matching guest:
```
DisplayName        : Jane External
UserPrincipalName  : jane_external.com#EXT#@yourtenant.onmicrosoft.com
Mail               : jane@external.com
Id                 : <guid>
UserType           : Guest
ExternalUserState  : Accepted
CreatedDateTime    : 2024-11-02 10:15:00
LatestSignIn       : Never
MatchReason        : Inactive > 180 days (last sign-in: )
------------------------------------------------------------
```

**2. Removal Report (`.csv`)** — only if `-Delete` is used. Columns:
`UserPrincipalName`, `DisplayName`, `UserType`, `MatchReason`, `Status`,
`Message`, `Date`.

**Status Definitions:**
- **Success:** Guest was successfully removed.
- **Skipped:** Matched the criteria but wasn't a Guest account, so it
  was left alone.
- **Failed:** See `Message` for the specific error (e.g. permission
  denied).

### Safety Notes

- **Review First:** Always run without `-Delete` first and review the
  `.txt` report before re-running with `-Delete`.
- **Soft Deletes:** Entra ID performs a soft delete. Removed accounts
  are recoverable for 30 days — see `Restore-DeletedGuestUsers.ps1`
  below.
- **Criteria are Absolute:** `-Force` *only* skips the confirmation
  prompt — it does not bypass the inactive/pending criteria itself.
  There's no flag to force-delete a guest that doesn't meet the rules.

---

## 3. Restore-DeletedGuestUsers.ps1

### Objective

Recovers guest accounts that were soft-deleted by
`Get-And-Remove-GuestUsers.ps1`, `Remove-StaleGuestUsers_v3.ps1`, or
`Remove-BulkGuestUsers.ps1`, within Entra ID's 30-day recovery window.
Since `Restore-MgDirectoryDeletedItem` requires the deleted object's
GUID (not an email or UPN), this script looks that GUID up for you from
a supplied email address, then restores it — individually or in bulk.

### How It Works

1. **Connects to Microsoft Graph:** Reuses an existing session if
   already signed in.
2. **Determines what to restore** via one of three paths:
   - `-ObjectId <guid>` — restores directly, no lookup needed.
   - `-Email <address>` — looks up that one guest in the recycle bin.
   - `-EmailListPath <file>` — bulk mode, one email per line.
   - No parameters — interactive prompt asks which mode you want.
3. **Looks up the deleted user's GUID:** pulls the full deleted-items
   list once via `Get-MgDirectoryDeletedItemAsUser`, then searches it
   using a wildcard/contains match against `Mail` and
   `UserPrincipalName` — not an exact match, since Entra ID can alter a
   guest's UPN/email on deletion (e.g. appending a random string or
   timestamp) to free the original address for reuse.
4. **Handles multiple matches:** if more than one deleted record
   matches, shows all candidates (display name, UPN, deletion date) and
   asks which one to restore.
5. **Confirms before restoring:** shows the full list of matched
   guests and asks for `Y/N` confirmation, unless `-Force` is passed.
6. **Restores and reports:** calls `Restore-MgDirectoryDeletedItem` for
   each confirmed match and writes a CSV report.

### Requirements

```powershell
Install-Module Microsoft.Graph.Users -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
```

### Permissions

| Scope | Purpose |
|---|---|
| `User.ReadWrite.All` (or `Directory.AccessAsUser.All`) | List deleted users, restore them |

**Entra ID Directory Role:** **User Administrator** or **Global
Administrator** — same requirement as the removal scripts.

**Time limit:** deleted users are only recoverable for **30 days**
after removal. After that, Entra ID permanently deletes them and no
tool — this script included — can recover them.

### Usage Examples

**Restore a single guest by email** (looks up the GUID for you):
```powershell
.\Restore-DeletedGuestUsers.ps1 -Email "alice@external.com"
```

**Restore directly if you already know the Object ID:**
```powershell
.\Restore-DeletedGuestUsers.ps1 -ObjectId "a9532b30-4edb-4b66-a3b0-6ac972a6065b"
```

**Bulk restore from a text file of emails, one per line:**
```powershell
.\Restore-DeletedGuestUsers.ps1 -EmailListPath "C:\Temp\restore-list.txt"
```

**No parameters — interactive prompt:**
```powershell
.\Restore-DeletedGuestUsers.ps1
```

**Skip the confirmation prompt (unattended use):**
```powershell
.\Restore-DeletedGuestUsers.ps1 -EmailListPath "C:\Temp\restore-list.txt" -Force
```

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-Email` | *(none)* | A single email address to look up and restore. |
| `-ObjectId` | *(none)* | Restore directly by known GUID, skipping the lookup. |
| `-EmailListPath` | *(none)* | Path to a `.txt` file of emails, one per line, for bulk restore. |
| `-ReportPath` | `GuestRestore_Report_<timestamp>.csv` | Where the restore CSV report is written. |
| `-Force` | *(off)* | Skip the `Y/N` confirmation prompt. |

If none of `-Email`, `-ObjectId`, or `-EmailListPath` are supplied, the
script prompts interactively for single-email or list-file mode.

### Output

**`GuestRestore_Report_<timestamp>.csv`** — columns: `Input`,
`ObjectId`, `Status`, `Message`, `Date`.

**Status Definitions:**
- **Success:** Guest restored successfully.
- **Failed:** Restore attempt failed — see `Message` (e.g. permission
  denied, outside recovery window).
- **NotFound:** No matching deleted user found for that email — likely
  outside the 30-day window, already restored, or the identifiers
  genuinely don't match.
- **Skipped:** Multiple deleted records matched the email and none was
  selected.

### Safety Notes

- **Matching is intentionally loose** (contains/wildcard, not exact) to
  account for Entra ID altering UPNs/emails on deletion — but every
  match is shown to you before anything is restored, so a broader match
  doesn't reduce safety, it just widens what gets surfaced for review.
- **Empty input is explicitly rejected**, not silently accepted — an
  unguarded blank search value would otherwise match every deleted user
  in the tenant via wildcard matching.
- All interactive prompts (mode selection, email entry, match
  selection, Y/N confirmation) re-prompt on invalid input rather than
  silently cancelling or proceeding on a typo.

---

## Shared Notes Across All Three Scripts

### Input file format

Any script that accepts a list of guests
(`-UserListPath`/`-EmailListPath`) uses the same plain-text convention:
```
alice@external.com
bob@partnercompany.com
carol@vendor.org
```
One UPN or email per line, no header, no commas/quotes, blank lines
ignored.

### Permissions matrix

| Script | Graph Scopes | Directory Role |
|---|---|---|
| `Get-GuestUserDetails.ps1` | `User.Read.All`, `AuditLog.Read.All` | Directory Readers / Reports Reader / Global Reader |
| `Get-And-Remove-GuestUsers.ps1` | `User.ReadWrite.All`, `AuditLog.Read.All` | User Administrator or Global Administrator |
| `Restore-DeletedGuestUsers.ps1` | `User.ReadWrite.All` | User Administrator or Global Administrator |

### Troubleshooting quick reference

| Symptom | Likely cause |
|---|---|
| Sign-in fields always empty | Tenant lacks Azure AD Premium P1/P2, or `AuditLog.Read.All` wasn't consented |
| Delete/restore fails with 403 despite successful sign-in | Signed-in account lacks the required directory role |
| "File not found" for a list-based script | Check the path — several scripts default to looking next to the script itself |
| Guest not found by `Restore-DeletedGuestUsers.ps1` | Outside the 30-day recovery window, already restored, or try `-ObjectId` directly if you have it |
| A guest looks "inactive" but you know they're active | Confirm the Premium license is active tenant-wide |

---
*FIRMUS INTERNAL*
