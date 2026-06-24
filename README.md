# Entra Tenant Governance TUI

A numbered, menu-driven **PowerShell console** (TUI) for the latest **Microsoft Entra ID
Identity Governance** management APIs, accessed through Microsoft Graph.

It shows the **logged-on account** at all times, lets you **format and export** any result
with simple commands, **installs its own prerequisites** from a menu, and can **update itself
from GitHub** with a single keystroke.

> Repository: <https://github.com/oguzhanf/tenantgovernance>
> Current version: **1.0.1** (2026-06-24) · Graph channels: **v1.0 + beta** · API reference verified **2026-06-24**

---

## Contents
- [What it does](#what-it-does)
- [Quick start](#quick-start)
- [Prerequisites (auto-installed)](#prerequisites-auto-installed)
- [Signing in](#signing-in)
- [Governance capabilities](#governance-capabilities)
- [Formatting & exporting output](#formatting--exporting-output)
- [Command reference](#command-reference)
- [Permissions / scopes](#permissions--scopes)
- [Versioning & self-update](#versioning--self-update)
- [Configuration file](#configuration-file)
- [Security notes](#security-notes)
- [Disclaimer](#disclaimer)

---

## What it does

```
================================================================================
                       ENTRA  TENANT  GOVERNANCE  -  TUI
                 Identity Governance management via Microsoft Graph
================================================================================
  Status : CONNECTED
  Account: admin@contoso.onmicrosoft.com
  Tenant : 00000000-0000-0000-0000-000000000000
  Auth   : Delegated   Scopes: 11   Graph: v1.0
  Prereqs: all satisfied
  Version: 1.0.0 (2026-06-24)   Graph: v1.0+beta   API ref: 2026-06-24
  Update : up to date (checked 09:14)
--------------------------------------------------------------------------------
  Governance
--------------------------------------------------------------------------------
     1. > Entitlement Management
     2. > Access Reviews
     3. > PIM - Directory Roles
     4. > PIM for Groups
     5. > Lifecycle Workflows
     6. > Terms of Use
     7. > App Consent
     8. > Tenant Overview
--------------------------------------------------------------------------------
    L. Login / Account     F. Format / settings     C. Custom Graph query
    P. Prerequisites/install  V. About / version     U. Update from GitHub
    ?. Help                                              Q. Quit
--------------------------------------------------------------------------------
```

- **Always-on status banner** — the signed-in account, tenant, auth type, granted scope
  count, prerequisite status, tool version, and update status. When you are **not** signed
  in, it says so and points you to **L** (Login).
- **Numbered menus** that expose every Identity Governance surface.
- **Direct Graph REST** via `Invoke-MgGraphRequest`, so it targets the **latest v1.0 and
  beta** endpoints without needing the full Graph SDK.
- **Formatted output** (Table / List / JSON / CSV / Grid) controlled by commands.
- **Export** any result to JSON/CSV/TXT or the clipboard.
- **Self-installing prerequisites** and **self-update** from this repo.

---

## Quick start

### Option A — clone and run
```powershell
git clone https://github.com/oguzhanf/tenantgovernance.git
cd tenantgovernance
pwsh ./EntraGovernanceTui.ps1
```

### Option B — one-line bootstrap (downloads the latest script and runs it)
```powershell
$f = Join-Path $env:TEMP 'EntraGovernanceTui.ps1'
Invoke-RestMethod 'https://raw.githubusercontent.com/oguzhanf/tenantgovernance/main/EntraGovernanceTui.ps1' -OutFile $f
pwsh -File $f
```

On first launch:
1. If `Microsoft.Graph.Authentication` is missing, the banner shows **`Prereqs: REQUIRED
   missing`** — press **`P`** then **`A`** to install everything automatically.
2. Press **`L`** → **`1`** to sign in interactively.
3. Pick a capability by number and explore.

> Requires **PowerShell 7+** (`pwsh`). If you start it in **Windows PowerShell 5.1**
> (e.g. `.\EntraGovernanceTui.ps1` in the classic console), it **automatically relaunches
> itself in `pwsh`** — or, if PowerShell 7 isn't installed, shows a one-line install command.

---

## Prerequisites (auto-installed)

The tool detects what it needs and can deploy it for you from the **Prerequisites** menu
(**`P`**). Nothing is installed without your action.

| Prerequisite | Required | Purpose | Auto-install |
|---|:---:|---|:---:|
| PowerShell 7.0+ | yes | Runtime | manual (`winget install Microsoft.PowerShell`) |
| PSGallery (trusted) | helper | Module source (no prompts) | yes |
| NuGet provider | helper | Module download provider | yes |
| **Microsoft.Graph.Authentication** | **yes** | `Connect-MgGraph`, `Invoke-MgGraphRequest` | yes |
| Microsoft.PowerShell.ConsoleGuiTools | optional | Interactive **Grid** output | yes |

In the **`P`** menu:
- **`A`** — install all *required* missing prerequisites (also fixes PSGallery/NuGet first).
- **`O`** — install all, including optional ones.
- **`<number>`** — install a single item.
- **`R`** — re-check.

All installs use `-Scope CurrentUser` (no admin rights required).

---

## Signing in

Press **`L`** for the Login / Account menu. **Interactive browser sign-in is the default**
(dynamic), and other modes are available:

| # | Mode | Use when |
|---|---|---|
| 1 | Interactive (browser) | Normal admin use (default) |
| 2 | Device code | Headless / remote / SSH sessions |
| 3 | App-only + certificate | Automation (ClientId + Tenant + thumbprint) |
| 4 | App-only + client secret | Automation (ClientId + Tenant + secret) |
| 5 | Managed identity | Running on an Azure-hosted machine |
| 6 | Sign out | — |
| 7 | Show full context | Inspect account + granted scopes |
| 8 | Set sign-in scopes | Override the requested delegated scopes |

You can also pre-select a tenant or scopes at launch:
```powershell
pwsh ./EntraGovernanceTui.ps1 -TenantId contoso.onmicrosoft.com
pwsh ./EntraGovernanceTui.ps1 -UseDeviceCode
```

> **Admin consent:** governance read scopes generally require admin consent the first time.
> An administrator can consent during interactive sign-in. App-only modes use the
> permissions configured on your app registration (note: **Terms of Use** is delegated-only
> and has no application permission).

---

## Governance capabilities

Every item issues a `GET` against the current Graph channel (v1.0 by default; switch with
`apiver beta`).

| Menu | Endpoints |
|---|---|
| **Entitlement Management** | access packages, catalogs, assignments, assignment requests, assignment policies, connected organizations, resource requests, settings |
| **Access Reviews** | review definitions, history definitions* |
| **PIM – Directory Roles** | role definitions, active assignments, assignment schedules / instances / requests, eligibility schedules / instances / requests |
| **PIM for Groups** | eligibility & assignment schedules / instances / requests (filter on group or principal) |
| **Lifecycle Workflows** | workflows, deleted workflows, templates, task definitions, custom task extensions, settings |
| **Terms of Use** | agreements |
| **App Consent** | app consent requests |
| **Tenant Overview** | organization, subscribed SKUs, directory roles |
| **Custom Graph query** (`C`) | any relative Graph path you type |

\* History definitions return the last 30 days and require a ReadWrite scope (see below).
**PIM for Groups** endpoints require a `$filter` on `groupId` or `principalId`; the tool
prompts you for it.

---

## Formatting & exporting output

After running any item you land in the **result view**, where commands control presentation:

```
format table|list|json|csv|grid   change the output format (alias: fmt; or just type the name)
cols id,displayName,status         choose table columns (blank = auto)
find <text>                        client-side filter the rows on screen
open <n>                           show full JSON for row n
top <n>                            server-side page size ($top)
maxitems <n>                       max rows fetched across pages
filter <odata>                     server-side $filter, e.g.  filter status eq 'active'
select / expand / orderby <...>    OData $select / $expand / $orderby
apiver auto|v1.0|beta              target stable or preview APIs
refresh (r)                        re-run with the current query options
export                             write JSON/CSV to the export folder (auto-named)
export <path>                      export to a specific .json / .csv / .txt file
export clip                        copy the result as JSON to the clipboard
b                                  back to the menu
```

---

## Command reference

Global commands work from the main menu prompt too:

| Command | Action |
|---|---|
| `L` / `login` | Login / Account menu |
| `who` | Show Graph context + granted scopes |
| `logout` | Sign out |
| `P` / `prereqs` | Prerequisites menu (install missing components) |
| `V` / `about` | About: version, API dates, update status |
| `U` / `update` | Check GitHub and apply an update if available |
| `F` | Output format / settings |
| `C` | Custom Graph query |
| `format <fmt>` | Set default output format |
| `apiver <auto\|v1.0\|beta>` | Set Graph channel |
| `?` / `help` | Help screen |
| `B` / `Q` | Back / Quit |

---

## Permissions / scopes

Default delegated scopes requested at sign-in (verified against Microsoft Learn on
**2026-06-24**):

| Scope | Covers |
|---|---|
| `Directory.Read.All` | Tenant overview; directory role defs/assignments |
| `EntitlementManagement.Read.All` | All Entitlement Management |
| `AccessReview.ReadWrite.All` | Access review **definitions + history** (history needs ReadWrite even to read) |
| `RoleManagement.Read.All` | PIM directory roles, schedules, instances |
| `RoleAssignmentSchedule.ReadWrite.Directory` | `roleAssignmentScheduleRequests` (GET requires ReadWrite) |
| `RoleEligibilitySchedule.ReadWrite.Directory` | `roleEligibilityScheduleRequests` (GET requires ReadWrite) |
| `PrivilegedEligibilitySchedule.Read.AzureADGroup` | PIM for Groups – eligibility |
| `PrivilegedAssignmentSchedule.Read.AzureADGroup` | PIM for Groups – assignment |
| `LifecycleWorkflows.Read.All` | Lifecycle Workflows |
| `Agreement.Read.All` | Terms of Use (delegated only) |
| `ConsentRequest.Read.All` | App consent requests |

Override the set anytime with **Login → 8 (Set sign-in scopes)**. Each menu item also knows
its own least-privilege scope and prints it if a call is denied.

---

## Versioning & self-update

- The tool version and the **API-reference-verified date** are shown in the banner and in
  **About** (`V`). Because the governance APIs move quickly, that date tells you when the
  endpoints and scopes were last checked against Microsoft Learn.
- On startup the tool quietly checks this repository and shows **`Update available`** in the
  banner when the published `EntraGovernanceTui.ps1` differs from your local copy (the
  comparison ignores line-ending/BOM differences).
- Press **`U`** to download, validate (it must parse cleanly), back up your current file to
  `EntraGovernanceTui.ps1.bak-<timestamp>`, and replace it — then optionally relaunch.
- **About** (`V`) also shows the remote commit date/SHA/message and the `version.json`
  manifest version, and lets you **set a different repo/branch** and **save** it.
- Skip the startup check with `-NoUpdateCheck`. Point at a fork with
  `-UpdateOwner <owner> -UpdateRepo <repo> -UpdateBranch <branch>`.

Maintainers: bump `$GovVersion` / `$GovReleaseDate` in the script **and** `version.json` when
publishing.

---

## Configuration file

Settings persist in `EntraGovernanceTui.config.json` next to the script (created when you
choose **Save config** in About). It can hold the update repo, default output format, API
channel, export directory, and a custom scope list. It is git-ignored.

```json
{
  "repo": { "owner": "oguzhanf", "name": "tenantgovernance", "branch": "main",
            "scriptPath": "EntraGovernanceTui.ps1", "manifestPath": "version.json" },
  "format": "Table",
  "apiVersion": "auto",
  "exportDir": "C:\\reports",
  "scopes": [ "EntitlementManagement.Read.All" ]
}
```

---

## Security notes

- The tool only issues **`GET`** requests — it reads governance data, it does not modify it.
  (A few endpoints require a `ReadWrite` scope even to read; see the scope table.)
- No secrets are stored in the script. App-only secrets you type are passed straight to
  `Connect-MgGraph` and never written to disk.
- Self-update only pulls from the configured GitHub repo over HTTPS and **validates** the
  download (PowerShell parse + sanity markers) and **backs up** your file before replacing it.
- Run only scripts you trust; review the source before granting tenant consent.

---

## Disclaimer

Provided "as is" under the [MIT License](LICENSE), with no warranty. Not an official
Microsoft product. Microsoft Graph and Entra are trademarks of Microsoft. API surface and
permissions can change; re-verify against
[Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/resources/identitygovernance-overview)
for production use.
