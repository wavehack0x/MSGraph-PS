# Guest User Lifecycle Cleanup | Get-And-Remove-GuestUsers.ps1

| | |
|---|---|
| **Author** | Waverly Chua |
| **Creation Date** | 5 Aug 2026 |
| **Version** | 1.0 |

## Objective

To maintain a secure and clean Microsoft Entra ID (Azure AD) environment
by systematically identifying and managing stale B2B guest accounts. The
**"Get-And-Remove-GuestUsers"** script provides an automated,
safety-first approach to guest lifecycle management. It audits your
tenant to find guests who are inactive or have lingering, unaccepted
invitations, logs these findings into a detailed plain-text report for
review, and — **only when explicitly instructed** — safely removes the
matching accounts.

## How It Works

1. **Connects to Microsoft Graph:** Reuses an existing session if you're
   already signed in.
2. **Retrieves Guest Users:** Fetches every guest in the tenant by
   default, or only the UPNs/emails listed in a text file via
   `-UserListPath`.
3. **Applies Filters:** Narrows down the list to guests matching your
   criteria:
   - **Inactive:** Invite Accepted, but no sign-in in more than
     `-InactiveDays` (default: 180 days).
   - **Stale Invite:** Invite still Pending Acceptance, sent more than
     `-PendingDays` ago (default: 90 days).
   - *(Note: Anyone matching neither condition is excluded
     entirely — not reported, not touched.)*
4. **Generates a Report:** Writes a plain `.txt` file with full details
   for every matching guest, including a `MatchReason`.
5. **Stops or Proceeds:**
   - **Default:** Stops here. No accounts are removed unless you pass
     the `-Delete` flag.
   - **If `-Delete` is passed:** Prompts for a `Y/N` confirmation
     (unless `-Force` is passed), then removes each matching guest one
     at a time. It skips any account that turns out not to actually be
     a Guest.
6. **Logs Removals:** Writes a CSV removal report tracking `Success`,
   `Failed`, or `Skipped` status per guest.

## Prerequisites & Permissions

### System Requirements

- **Module:**
  ```powershell
  Install-Module Microsoft.Graph.Users -Scope CurrentUser
  ```
- **Licensing:** Azure AD Premium P1 or P2 license. This is **required**
  for the `SignInActivity` data that the inactivity check relies on.
  Without it, sign-in fields return empty, and every guest may falsely
  appear "inactive".

### Required Permissions

You must have both the correct Graph API scope **and** the appropriate
Entra ID directory role.

**1. Graph API Scopes:**

| Scope | Purpose |
|---|---|
| `User.ReadWrite.All` | Read guest profiles, and delete them if `-Delete` is used |
| `AuditLog.Read.All` | Read `SignInActivity` for the inactivity check and text report |

> Note: an earlier draft of this table listed `User.Read` and
> `User.Write` as separate scopes. Microsoft Graph doesn't have a
> standalone `User.Write` scope — the script requests the single
> combined scope `User.ReadWrite.All`, which covers both reading and
> writing/deleting guest profiles.

**2. Entra ID Directory Role:**

Having the API scopes above is not enough to delete users. The
signed-in account also needs one of these roles:
- **User Administrator**
- **Global Administrator**

*(Without one of these roles, the `-Delete` command will fail with a
permission error, even if sign-in and scope consent succeed.)*

## Usage Examples

### Report only (Default) — Always start here

Writes a text file listing every guest matching the default criteria
(Inactive > 180 days OR Pending > 90 days). No accounts are touched.

```powershell
.\Get-And-Remove-GuestUsers.ps1
```

### Custom thresholds

Adjust the inactivity and pending invite thresholds.

```powershell
.\Get-And-Remove-GuestUsers.ps1 -InactiveDays 90 -PendingDays 30
```

### Scope to a specific list of guests

The inactive/pending criteria still applies on top of this list. A
listed guest who is actually still active or recently invited will be
excluded from the report.

```powershell
.\Get-And-Remove-GuestUsers.ps1 -UserListPath "C:\Temp\guests.txt"
```

### Report + Delete, with confirmation prompt

```powershell
.\Get-And-Remove-GuestUsers.ps1 -Delete
```

### Report + Delete, fully unattended (no prompt)

Use this for scheduled/automated runs (Azure Automation, Power
Automate, Task Scheduler).

```powershell
.\Get-And-Remove-GuestUsers.ps1 -Delete -Force
```

> **Note on Automation:** This script uses interactive `Connect-MgGraph`
> sign-in by default. For a truly unattended scheduled run, you must
> switch authentication to app-only (certificate or client secret)
> inside the script, otherwise it will hang waiting for a login prompt.

## Parameters

| Parameter | Default | Description |
|---|---|---|
| `-UserListPath` | *(none — all guests)* | Path to a `.txt` file of specific UPNs/emails to scope to, one per line. |
| `-DetailsPath` | `guest-details_<timestamp>.txt` | Where the txt detail report is written (saves to script folder). |
| `-ReportPath` | `GuestUserRemoval_Report_<timestamp>.csv` | Where the removal CSV report is written (only created if `-Delete` is used). |
| `-InactiveDays` | `180` | Days since last sign-in before an Accepted guest counts as inactive. |
| `-PendingDays` | `90` | Number of days after an invitation is sent before a pending guest is considered stale. |
| `-Delete` | *(off)* | Remove matching guests. Without this, the script is report-only. |
| `-Force` | *(off)* | Skip the `Y/N` confirmation prompt when `-Delete` is used. |

**Guest UPN Sample:** `johndoe@gmail.com#EXT#@yourtenant.onmicrosoft.com`

## Output Files

### 1. Detail Report (`.txt`)

Generated every time the script runs. Creates one block per matching
guest:

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

### 2. Removal Report (`.csv`)

Generated **only** if `-Delete` is used. Contains the following
columns: `UserPrincipalName`, `DisplayName`, `UserType`, `MatchReason`,
`Status`, `Message`, `Date`.

**Status Definitions:**
- **Success:** Guest was successfully removed.
- **Skipped:** Matched the criteria but wasn't a Guest account, so it
  was left alone.
- **Failed:** See the `Message` column for the specific error (e.g.,
  permission denied).

## Safety Notes

- **Review First:** Always run without `-Delete` first and review the
  `.txt` report before re-running with the `-Delete` switch.
- **Soft Deletes:** Entra ID performs a soft delete. Removed accounts
  are recoverable for 30 days via `Restore-MgDirectoryDeletedItem` if
  something gets caught by the criteria unexpectedly.
- **Criteria are Absolute:** The `-Force` parameter *only* skips the
  confirmation prompt — it does not bypass the inactive/pending
  criteria itself. There is currently no flag in this script to
  force-delete a guest that does not meet the inactivity rules.

---
