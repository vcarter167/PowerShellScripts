<div align="center">

# ⚡ PowerShell Scripts
### M365 · Exchange Online · Entra ID Automation

**Stop clicking through admin portals. Start running scripts.**

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?style=for-the-badge&logo=powershell&logoColor=white)](https://docs.microsoft.com/en-us/powershell/)
[![Microsoft 365](https://img.shields.io/badge/Microsoft_365-Exchange_Online-D83B01?style=for-the-badge&logo=microsoft&logoColor=white)](https://learn.microsoft.com/en-us/exchange/exchange-online)
[![Entra ID](https://img.shields.io/badge/Entra_ID-Azure_AD-0078D4?style=for-the-badge&logo=microsoft-azure&logoColor=white)](https://learn.microsoft.com/en-us/entra/identity/)
[![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](LICENSE)

</div>

---

## 🧠 The Problem This Solves

Every M365 admin knows the grind:

- A new vendor list comes in — 40 external contacts need to be created *and* added to a distribution group
- A department reorg just happened — 80 user profiles need their title, department, and office updated in Entra ID
- A new team is stood up — 25 internal users need access to a shared mailbox before 9am Monday

Done manually through the admin portal, these tasks are **slow, error-prone, and completely unscalable.** Miss a step, fat-finger a field, or lose track of where you are in the list — and you're doing it over again.

This repo is a collection of PowerShell scripts built to eliminate that entirely. **Prepare a CSV. Run the script. Done.**

---

## 📁 Scripts

---

### 1. `Add-ContactsToDistributionList.ps1`
> *Create external mail contacts and add them to a distribution group in one pass*

**The scenario:** You've received a list of external partners, vendors, or organization contacts that need to be reachable via a shared distribution list. Normally this is two separate manual workflows — create each contact in Exchange, then add them to the group.

**What the script does:**
- Prompts for your admin email and connects to Exchange Online via MFA browser authentication (`-DisableWAM`)
- Prompts whether you are adding a **single contact** or **bulk importing** from a file
- Prompts for the **target distribution list** name or email address and validates it exists before proceeding
- **Single mode:** prompts for the contact's full name and external email address with inline validation
- **Bulk mode:** accepts a `.csv` or `.xlsx`/`.xls` file — validates required columns, skips blank or malformed rows, and prints a summary on completion
- For each contact: checks if the mail contact already exists — if not, creates it via `New-MailContact` with parsed first/last name
- Checks if the contact is already a member of the target distribution list — if not, adds them via `Add-DistributionGroupMember`
- Skips duplicates gracefully and logs every action to the console
- Always disconnects from Exchange Online on completion, even if an error occurs

**Time saved:** What takes 2–3 minutes per contact manually runs in seconds per record at scale.

---

### 2. `Add-UsersToDistributionList.ps1`
> *Bulk-add internal M365 users to a distribution group from a CSV*

**The scenario:** A new distribution list is being created — or an existing one needs to be populated from an HR export or onboarding roster. Adding 50 users one at a time through the Exchange Admin Center is not the move.

**What the script does:**
- Prompts for your admin email and connects to Exchange Online via MFA browser authentication (`-DisableWAM`)
- Prompts whether you are adding a **single user** or **bulk importing** from a file
- Prompts for the **target distribution list** name or email address and validates it exists before proceeding
- **Single mode:** prompts for the user's full name and email address with inline validation
- **Bulk mode:** accepts a `.csv` or `.xlsx`/`.xls` file — validates required columns, skips blank or malformed rows, and prints a summary on completion
- For each user: checks for existing membership before attempting the add — no duplicate errors, no noise
- Adds via `Add-DistributionGroupMember` with **color-coded console output**:
  - 🟢 **Green** — user successfully added
  - 🟡 **Yellow** — user already a member, skipped
  - 🔴 **Red** — error encountered
- Always disconnects from Exchange Online on completion, even if an error occurs

---

### 3. `Add-UsersToSharedMailbox.ps1`
> *Bulk shared mailbox access provisioning via console prompts*

**The scenario:** A new department inbox or project mailbox is being stood up and a whole team needs Full Access. You need this done fast, and you may need to hand it off to someone else to run.

**What the script does:**
- Prompts for your admin email and connects to Exchange Online via MFA browser authentication (`-DisableWAM`)
- Prompts for the **shared mailbox email address** and validates it exists in Exchange before proceeding
- Prompts whether you are adding a **single user** or **bulk importing** from a file
- **Single mode:** prompts for the user's full name and email address with inline validation
- **Bulk mode:** accepts a `.csv` or `.xlsx`/`.xls` file — validates required columns, skips blank or malformed rows, and prints a summary on completion
- Loops through each user and grants **Full Access** via `Add-MailboxPermission`
- Suppresses known benign `-Member` parameter warnings while surfacing real errors in red
- Always disconnects from Exchange Online on completion, even if an error occurs

---

### 4. `Get-UserDistributionGroups.ps1`
> *Look up all distribution lists a user belongs to — single or bulk*

**The scenario:** An offboarding request comes in, or you need to audit group memberships before a reorg. Finding every distribution list a user belongs to manually through the Exchange Admin Center means searching one group at a time.

**What the script does:**
- Prompts for your admin email and connects to Exchange Online via MFA browser authentication (`-DisableWAM`)
- Prompts whether you are looking up a **single user** or **bulk looking up** from a file
- Pre-fetches all distribution groups once before processing — significantly faster than querying per user on large tenants
- **Single mode:** prompts for the user's email address and prints all matched groups in a formatted table
- **Bulk mode:** accepts a `.csv` or `.xlsx`/`.xls` file with an `Email` column (`Name` optional) — processes each user and collects all results
- At the end of a bulk run, offers to **export results to CSV** for reporting or compliance purposes
- Always disconnects from Exchange Online on completion, even if an error occurs

---

### 5. `Update-EntraIDUserProfiles.ps1`
> *Bulk update job titles, departments, and office locations in Entra ID from a CSV*

**The scenario:** Post-reorg. Annual HR data sync. Title changes after a performance cycle. Whatever the trigger — you have a spreadsheet of users with updated attributes and a directory that doesn't reflect reality yet. Touching each profile manually in the Entra admin portal for 60+ people is hours of work.

**What the script does:**
- Sets `ExecutionPolicy Bypass` and connects to Entra ID via `Connect-AzureAD`
- Reads user data from a CSV keyed on **Object ID** (the most reliable identifier)
- For each user, simultaneously updates three attributes:
  - **Job Title** via `Set-AzureADUser -JobTitle`
  - **Department** — auto-strips the first 3 characters from the raw value before writing (handles HRIS exports from systems like Workday or ADP that prepend numeric department codes, e.g. `001-Engineering` → `Engineering`)
  - **Office Location** via `-PhysicalDeliveryOfficeName`
- Writes a per-user confirmation to the console on completion

> ⚠️ This script uses `Set-ExecutionPolicy Bypass`. Verify this is permitted under your organization's security policy before executing in production.

```powershell
# Key variable to configure before running:
$userData = Import-Csv -Path "C:\Path\To\File.csv"
```

---

## 📄 CSV Reference

> Headers must match exactly — including capitalization and spacing.

### Exchange Scripts (Scripts 1, 2, 3, 4)

```csv
Name,Email
Jane Smith,jsmith@example.com
John Doe,jdoe@externalpartner.com
```

| Column | Required | Notes |
|--------|----------|-------|
| `Name` | ✅ | Full display name. For external contacts, must be `First Last` — script splits on space for `New-MailContact`. Optional in Script 4 bulk mode. |
| `Email` | ✅ | Primary SMTP address |

---

### Entra ID Script (Script 5)

```csv
Object Id,Display name,Title,Department,Office
a1b2c3d4-xxxx-xxxx-xxxx-xxxxxxxxxxxx,Jane Smith,Senior Engineer,001-Engineering,New York
```

| Column | Required | Notes |
|--------|----------|-------|
| `Object Id` | ✅ | Entra ID Object ID — used as the `-ObjectId` targeting identifier |
| `Display name` | ✅ | Used for console confirmation output only — not written to directory |
| `Title` | ✅ | Job title to set |
| `Department` | ✅ | Script auto-strips first 3 characters — handles HRIS numeric prefixes |
| `Office` | ✅ | Physical office location |

---

## ⚙️ Prerequisites

| Requirement | Scripts | Details |
|-------------|---------|---------|
| **PowerShell** | All | Version 5.1+ or PowerShell 7+ |
| `ExchangeOnlineManagement` | 1, 2, 3, 4 | Exchange Online connectivity |
| `ImportExcel` | 1, 2, 3, 4 | Required for `.xlsx`/`.xls` bulk import support — auto-installed by scripts |
| `AzureAD` | 5 | Entra ID (Azure AD) connectivity |
| **Exchange Admin** or **Global Admin** | 1, 2, 3, 4 | Required to manage distribution groups and mailboxes |
| **User Administrator** or **Global Admin** | 5 | Required to write user profile attributes in Entra ID |

**Install modules:**

```powershell
# Exchange Online Management
Install-Module -Name ExchangeOnlineManagement -Force

# ImportExcel (for .xlsx support)
Install-Module -Name ImportExcel -Force

# Entra ID (Azure AD)
Install-Module -Name AzureAD -Force
```

---

## 🚀 Quick Start

```powershell
# 1. Clone the repo
git clone https://github.com/vcarter167/PowerShellScripts.git
cd PowerShellScripts

# 2. Install required modules (if not already installed)
Install-Module -Name ExchangeOnlineManagement -Force
Install-Module -Name ImportExcel -Force
Install-Module -Name AzureAD -Force

# 3. Run the script you need — all prompts are interactive, nothing to edit
.\Add-ContactsToDistributionList.ps1

# 4. A browser window will open — sign in and complete MFA when prompted
```

All Exchange scripts will prompt you for your admin email and authenticate via browser MFA using `-DisableWAM`. **No credentials are stored or hardcoded.**

---

## ⚠️ Known Behavior

When `Get-DistributionGroupMember` is called with a `-Member` filter in certain Exchange Online module versions, a non-terminating parameter error may appear in the console. **This does not interrupt execution** — membership additions complete successfully. This is a known module quirk and does not affect output integrity.

---

## 🗺️ Roadmap

| Status | Script | Description |
|--------|--------|-------------|
| ✅ Done | `Add-ContactsToDistributionList.ps1` | Bulk create external contacts + add to distro list |
| ✅ Done | `Add-UsersToDistributionList.ps1` | Bulk add internal users to distribution list |
| ✅ Done | `Add-UsersToSharedMailbox.ps1` | Bulk Full Access provisioning to shared mailbox |
| ✅ Done | `Get-UserDistributionGroups.ps1` | Look up all distribution lists a user belongs to |
| ✅ Done | `Update-EntraIDUserProfiles.ps1` | Bulk update title, department, office via Entra ID |
| 🔜 Planned | `Remove-UsersFromDistributionList.ps1` | Bulk offboarding from distribution groups |
| 🔜 Planned | `Audit-DistributionListMembers.ps1` | Export full membership to CSV for compliance review |
| 🔜 Planned | Graph API Migration | Replace legacy `AzureAD` + Exchange cmdlets with Microsoft Graph |

---

## 👤 Author

<div align="center">

**Vince Carter**
Systems Administrator | M365 · Entra ID · IT Automation

[![LinkedIn](https://img.shields.io/badge/LinkedIn-vcarter167-0077B5?style=flat-square&logo=linkedin)](https://www.linkedin.com/in/vcarter167)
[![GitHub](https://img.shields.io/badge/GitHub-vcarter167-181717?style=flat-square&logo=github)](https://github.com/vcarter167)
[![Email](https://img.shields.io/badge/Email-vcarter167%40gmail.com-D14836?style=flat-square&logo=gmail)](mailto:vcarter167@gmail.com)

*"Automate everything. Document everything. Own your infrastructure."*

</div>
