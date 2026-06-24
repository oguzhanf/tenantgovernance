#Requires -Version 7.0
<#
.SYNOPSIS
    Entra Tenant Governance TUI — a numbered, menu-driven console tool for the
    Microsoft Entra ID (Identity) Governance management APIs via Microsoft Graph.

.DESCRIPTION
    Exposes Entitlement Management, Access Reviews, Privileged Identity Management
    (directory roles and groups), Lifecycle Workflows, Terms of Use and App Consent
    governance surfaces. Output formatting and export are controlled with commands.

    Authentication uses the Microsoft.Graph.Authentication module. All data calls go
    through Invoke-MgGraphRequest (raw Graph REST) so the latest v1.0 and beta
    endpoints can be targeted without the full Graph SDK.

.PARAMETER TenantId
    Optional tenant id (GUID) or domain to sign in to.

.PARAMETER Scopes
    Optional override of the delegated scopes requested at interactive/device login.

.PARAMETER UseDeviceCode
    Start with device-code sign-in instead of the interactive browser flow.

.PARAMETER LoadOnly
    Dot-source the script to define functions WITHOUT launching the TUI (for testing).
    Example:  . .\EntraGovernanceTui.ps1 -LoadOnly

.EXAMPLE
    .\EntraGovernanceTui.ps1
    Launches the TUI; choose "L" to sign in interactively.

.EXAMPLE
    .\EntraGovernanceTui.ps1 -UseDeviceCode -TenantId contoso.onmicrosoft.com
#>
[CmdletBinding()]
param(
    [string]   $TenantId,
    [string[]] $Scopes,
    [switch]   $UseDeviceCode,
    [string]   $UpdateOwner,
    [string]   $UpdateRepo,
    [string]   $UpdateBranch,
    [switch]   $NoUpdateCheck,
    [switch]   $LoadOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region ------------------------------------------------- Version & update source

# Tool identity. Bump $GovVersion + $GovReleaseDate on every published change and
# mirror them in version.json so the in-tool update check can show "vX -> vY".
$script:GovVersion         = '1.0.0'
$script:GovReleaseDate     = '2026-06-24'
# Microsoft Graph channels this tool targets, and the date the governance
# endpoints + permission scopes were last verified against learn.microsoft.com.
$script:GovGraphChannels   = @('v1.0', 'beta')
$script:GovApiVerifiedDate = '2026-06-24'
# Known upstream deprecation worth surfacing to admins.
$script:GovApiNotes        = 'Legacy beta PIM (privilegedRoles*) stops returning data 2026-10-28; this tool uses the current iteration-3 roleManagement APIs.'

# Absolute path to THIS script (used by the self-update feature to replace itself).
$script:GovScriptPath = if ($PSCommandPath) { $PSCommandPath } else { Join-Path (Get-Location).Path 'EntraGovernanceTui.ps1' }

# GitHub source for self-update. Overridable via params or the saved config file.
$script:GovRepo = [ordered]@{
    Owner        = 'oguzhanf'
    Name         = 'tenantgovernance'
    Branch       = 'main'
    ScriptPath   = 'EntraGovernanceTui.ps1'
    ManifestPath = 'version.json'
}
if ($UpdateOwner)  { $script:GovRepo.Owner  = $UpdateOwner }
if ($UpdateRepo)   { $script:GovRepo.Name   = $UpdateRepo }
if ($UpdateBranch) { $script:GovRepo.Branch = $UpdateBranch }

$script:GovUpdate        = $null    # cached update-check result
$script:GovRequestExit   = $false   # set true to break the main loop (e.g. after relaunch)
$script:GovNoUpdateCheck = [bool]$NoUpdateCheck

#endregion


#region ---------------------------------------------------------------- State

function Initialize-GovState {
    [CmdletBinding()]
    param()

    $script:Gov = [ordered]@{
        GraphBaseUrl  = 'https://graph.microsoft.com'
        # Default delegated (read) scopes covering every governance surface below.
        # Default delegated scopes. Read-oriented, but a few endpoints require a
        # ReadWrite scope even for GET (verified vs learn.microsoft.com 2026-06-24):
        #   - accessReviews/historyDefinitions      -> AccessReview.ReadWrite.All
        #   - role*ScheduleRequests (GET)           -> Role*Schedule.ReadWrite.Directory
        # An admin must consent to these the first time. Override via the Login menu.
        DefaultScopes = @(
            'Directory.Read.All'                            # tenant overview + role defs/assignments
            'EntitlementManagement.Read.All'                # all entitlement management
            'AccessReview.ReadWrite.All'                    # access review definitions + historyDefinitions
            'RoleManagement.Read.All'                       # PIM directory roles/schedules/instances
            'RoleAssignmentSchedule.ReadWrite.Directory'    # roleAssignmentScheduleRequests (GET needs RW)
            'RoleEligibilitySchedule.ReadWrite.Directory'   # roleEligibilityScheduleRequests (GET needs RW)
            'PrivilegedEligibilitySchedule.Read.AzureADGroup' # PIM for groups - eligibility
            'PrivilegedAssignmentSchedule.Read.AzureADGroup'  # PIM for groups - assignment
            'LifecycleWorkflows.Read.All'                   # lifecycle workflows
            'Agreement.Read.All'                            # terms of use (delegated only)
            'ConsentRequest.Read.All'                       # app consent requests
        )
        Scopes        = $null     # effective scopes (override or default)
        Format        = 'Table'   # Table | List | Json | Csv | Grid
        ApiVersion    = 'auto'    # auto | v1.0 | beta
        Top           = 100       # server-side $top per page (0 = omit)
        MaxItems      = 200       # client-side ceiling across pages
        Filter        = $null     # $filter override (per action run)
        Select        = $null     # $select override
        Expand        = $null     # $expand override
        OrderBy       = $null     # $orderby override
        LastResult    = @()       # raw objects from last action
        LastView      = @()       # possibly client-filtered view
        LastAction    = $null     # the action node last run
        ExportDir     = (Get-Location).Path
        PrereqStatus  = $null
        PrereqQuick   = $null
        Menu          = $null
    }
    $script:Gov.Scopes = if ($script:GovScopeOverride) { $script:GovScopeOverride } else { $script:Gov.DefaultScopes }
    $script:Gov.Menu   = Get-GovMenuTree
}

#endregion

#region -------------------------------------------------------- Prerequisites

# Static definitions. Each prereq exposes a Check (returns status) and an Install
# script block. Order matters: PSGallery + NuGet must precede module installs.
function Get-GovPrereqDefinitions {
    [CmdletBinding()] param()
    @(
        [ordered]@{
            Key='pwsh'; Name='PowerShell 7.0+'; Required=$true; AutoInstall=$false; Runtime=$true; Helper=$false
            Detail='Modern PowerShell runtime required by this tool.'
            Check = {
                $ok = $PSVersionTable.PSVersion.Major -ge 7
                [ordered]@{ Installed=$ok; Version="$($PSVersionTable.PSVersion)"; Note=($ok ? 'OK' : 'Upgrade required') }
            }
            Install = {
                Write-Host '  PowerShell 7 cannot install itself from within itself.' -ForegroundColor Yellow
                Write-Host '  Install with:  winget install --id Microsoft.PowerShell -e' -ForegroundColor Yellow
                Write-Host '  Then re-launch this tool in "pwsh".' -ForegroundColor Yellow
            }
        }
        [ordered]@{
            Key='psgallery'; Name='PSGallery repository (trusted)'; Required=$false; AutoInstall=$true; Runtime=$false; Helper=$true
            Detail='Package source used to download the modules below (trusted = no prompts).'
            Check = {
                $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
                if (-not $repo) { return [ordered]@{ Installed=$false; Version='-'; Note='Not registered' } }
                $trusted = $repo.InstallationPolicy -eq 'Trusted'
                [ordered]@{ Installed=$trusted; Version=$repo.InstallationPolicy; Note=($trusted ? 'Trusted' : 'Untrusted (would prompt)') }
            }
            Install = {
                if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
                    Register-PSRepository -Default -ErrorAction Stop
                }
                Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
            }
        }
        [ordered]@{
            Key='nuget'; Name='NuGet package provider'; Required=$false; AutoInstall=$true; Runtime=$false; Helper=$true
            Detail='Provider PowerShellGet uses to acquire modules from PSGallery.'
            Check = {
                # NOTE: never use -ListAvailable here; it scans disk and can take ~90s.
                $p = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue |
                     Sort-Object Version -Descending | Select-Object -First 1
                if ($p) { [ordered]@{ Installed=$true; Version="$($p.Version)"; Note='OK' } }
                else    { [ordered]@{ Installed=$false; Version='-'; Note='Needed only to install modules' } }
            }
            Install = {
                Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
            }
        }
        [ordered]@{
            Key='graph-auth'; Name='Microsoft.Graph.Authentication'; Required=$true; AutoInstall=$true; Runtime=$true; Helper=$false; MinVersion='2.0.0'
            Detail='Core dependency: Connect-MgGraph, Get-MgContext, Invoke-MgGraphRequest.'
            Check = {
                $m = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication |
                     Sort-Object Version -Descending | Select-Object -First 1
                if ($m) { [ordered]@{ Installed=$true; Version="$($m.Version)"; Note='OK' } }
                else    { [ordered]@{ Installed=$false; Version='-'; Note='Required for sign-in & API calls' } }
            }
            Install = {
                Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            }
        }
        [ordered]@{
            Key='consolegui'; Name='Microsoft.PowerShell.ConsoleGuiTools'; Required=$false; AutoInstall=$true; Runtime=$true; Helper=$false
            Detail='Optional: enables the interactive "Grid" output format (Out-ConsoleGridView).'
            Check = {
                $m = Get-Module -ListAvailable -Name Microsoft.PowerShell.ConsoleGuiTools |
                     Sort-Object Version -Descending | Select-Object -First 1
                if ($m) { [ordered]@{ Installed=$true; Version="$($m.Version)"; Note='OK' } }
                else    { [ordered]@{ Installed=$false; Version='-'; Note='Optional (Grid view)' } }
            }
            Install = {
                Install-Module -Name Microsoft.PowerShell.ConsoleGuiTools -Scope CurrentUser -Force -ErrorAction Stop
            }
        }
    )
}

# Fast check used by the banner/startup: only the runtime-relevant items (no slow
# package-provider / repository calls). The installer helpers (PSGallery, NuGet)
# are evaluated lazily in the full Prerequisites menu.
function Update-GovPrereqQuick {
    [CmdletBinding()] param()
    $reqMiss = 0; $optMiss = 0
    foreach ($d in (Get-GovPrereqDefinitions | Where-Object { $_.Runtime })) {
        $res = & $d.Check
        if (-not $res.Installed) { if ($d.Required) { $reqMiss++ } else { $optMiss++ } }
    }
    $script:Gov.PrereqQuick = [pscustomobject]@{ RequiredMissing=$reqMiss; OptionalMissing=$optMiss; Checked=(Get-Date) }
    return $script:Gov.PrereqQuick
}

function Update-GovPrereqStatus {
    [CmdletBinding()] param()
    $list = foreach ($def in Get-GovPrereqDefinitions) {
        $res = & $def.Check
        [pscustomobject]@{
            Key=$def.Key; Name=$def.Name; Required=$def.Required; AutoInstall=$def.AutoInstall
            Runtime=$def.Runtime; Helper=$def.Helper
            Detail=$def.Detail; Installed=[bool]$res.Installed; Version=$res.Version; Note=$res.Note
            Install=$def.Install
        }
    }
    $script:Gov.PrereqStatus = @($list)
    # keep the quick summary in sync after a full scan
    $rm = @($list | Where-Object { $_.Runtime -and $_.Required -and -not $_.Installed }).Count
    $om = @($list | Where-Object { $_.Runtime -and -not $_.Required -and -not $_.Installed }).Count
    $script:Gov.PrereqQuick = [pscustomobject]@{ RequiredMissing=$rm; OptionalMissing=$om; Checked=(Get-Date) }
    return $script:Gov.PrereqStatus
}

function Show-GovPrereqBannerLine {
    [CmdletBinding()] param()
    $q = $script:Gov.PrereqQuick
    if (-not $q) { $q = Update-GovPrereqQuick }
    Write-Host '  Prereqs: ' -NoNewline
    if ($q.RequiredMissing -gt 0) {
        Write-Host ("{0} REQUIRED missing - press 'P' to install" -f $q.RequiredMissing) -ForegroundColor Red
    }
    elseif ($q.OptionalMissing -gt 0) {
        Write-Host ("required OK ({0} optional available - 'P')" -f $q.OptionalMissing) -ForegroundColor Yellow
    }
    else {
        Write-Host 'all satisfied' -ForegroundColor Green
    }
}

function Install-GovPrereqItem {
    [CmdletBinding()] param([Parameter(Mandatory)] $Item)
    Write-Host ("  Installing: {0} ..." -f $Item.Name) -ForegroundColor Cyan
    try {
        & $Item.Install
        Write-Host ("  Done: {0}" -f $Item.Name) -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host ("  Failed: {0} - {1}" -f $Item.Name, $_.Exception.Message) -ForegroundColor Red
        return $false
    }
}

function Install-GovMissingPrereqs {
    [CmdletBinding()] param([switch] $IncludeOptional)
    $st = Update-GovPrereqStatus
    # Modules to install (required, plus optional when asked).
    $modules = @($st | Where-Object {
        -not $_.Installed -and $_.AutoInstall -and -not $_.Helper -and ($_.Required -or $IncludeOptional)
    })
    if ($modules.Count -eq 0) { Write-Host '  Nothing to install - all set.' -ForegroundColor Green; return }
    # Installer helpers (PSGallery trust, NuGet) first so module installs are seamless.
    $helpers = @($st | Where-Object { -not $_.Installed -and $_.AutoInstall -and $_.Helper })
    foreach ($t in ($helpers + $modules)) { [void](Install-GovPrereqItem -Item $t) }
    [void](Update-GovPrereqStatus)
    if ($script:Gov.PrereqQuick.RequiredMissing -eq 0) { Write-Host '  All required prerequisites are now satisfied.' -ForegroundColor Green }
    else { Write-Host ("  {0} required prerequisite(s) still missing." -f $script:Gov.PrereqQuick.RequiredMissing) -ForegroundColor Yellow }
}

function Show-GovPrereqMenu {
    [CmdletBinding()] param()
    Write-Host '  Scanning prerequisites...' -ForegroundColor DarkGray
    $st = Update-GovPrereqStatus
    while ($true) {
        Clear-GovHost
        Show-GovBanner
        Write-Host '  PREREQUISITES  (deploy everything this tool needs on this machine)' -ForegroundColor White
        Write-GovRule
        $i = 1
        foreach ($p in $st) {
            $statusText = if ($p.Installed) { 'OK      ' } elseif ($p.Required) { 'MISSING ' } elseif ($p.Helper) { 'HELPER  ' } else { 'OPTIONAL' }
            $color      = if ($p.Installed) { 'Green'   } elseif ($p.Required) { 'Red'      } else { 'Yellow' }
            Write-Host '   ' -NoNewline
            Write-Host ('{0,2}.' -f $i) -ForegroundColor Yellow -NoNewline
            Write-Host ' [' -NoNewline
            Write-Host $statusText -ForegroundColor $color -NoNewline
            Write-Host ('] {0,-36} {1}' -f $p.Name, $p.Version)
            Write-Host ('          {0}' -f $p.Detail) -ForegroundColor DarkGray
            $i++
        }
        Write-GovRule
        Write-Host '   A. Install all REQUIRED missing       O. Install all (incl. optional)'
        Write-Host '   <number>. Install one item            R. Re-check        B. Back'
        Write-GovRule
        $sel = (Read-Host '   prereq>').Trim().ToLower()
        switch ($sel) {
            'a' { Install-GovMissingPrereqs;                 $st = Update-GovPrereqStatus; Read-Host '   Press Enter to continue' | Out-Null }
            'o' { Install-GovMissingPrereqs -IncludeOptional; $st = Update-GovPrereqStatus; Read-Host '   Press Enter to continue' | Out-Null }
            'r' { Write-Host '   Re-scanning...' -ForegroundColor DarkGray; $st = Update-GovPrereqStatus }
            'b' { return }
            ''  { }
            default {
                if ($sel -match '^\d+$') {
                    $idx = [int]$sel - 1
                    if ($idx -ge 0 -and $idx -lt $st.Count) {
                        $item = $st[$idx]
                        if ($item.Installed)            { Write-Host '   Already installed.' -ForegroundColor Green }
                        elseif (-not $item.AutoInstall) { & $item.Install }
                        else                            { [void](Install-GovPrereqItem -Item $item) }
                        $st = Update-GovPrereqStatus
                        Read-Host '   Press Enter to continue' | Out-Null
                    }
                    else { Write-Host '   Invalid number.' -ForegroundColor Yellow; Start-Sleep -Milliseconds 700 }
                }
                else { Write-Host '   Unknown option.' -ForegroundColor Yellow; Start-Sleep -Milliseconds 700 }
            }
        }
    }
}

#endregion

#region --------------------------------------------- Versioning / config / update

function Get-GovProp {
    param($Obj, [string] $Name)
    if ($null -ne $Obj -and ($Obj.PSObject.Properties.Name -contains $Name)) { return $Obj.$Name }
    return $null
}

function Get-GovConfigPath {
    $dir = if ($script:GovScriptPath) { Split-Path -Parent $script:GovScriptPath } else { (Get-Location).Path }
    Join-Path $dir 'EntraGovernanceTui.config.json'
}

function Import-GovConfig {
    [CmdletBinding()] param()
    $p = Get-GovConfigPath
    if (-not (Test-Path $p)) { return }
    try {
        $cfg  = Get-Content -Path $p -Raw | ConvertFrom-Json
        $repo = Get-GovProp $cfg 'repo'
        if ($repo) {
            foreach ($k in 'owner','name','branch','scriptPath','manifestPath') {
                $v = Get-GovProp $repo $k
                if ($v) {
                    switch ($k) {
                        'owner'        { $script:GovRepo.Owner        = $v }
                        'name'         { $script:GovRepo.Name         = $v }
                        'branch'       { $script:GovRepo.Branch       = $v }
                        'scriptPath'   { $script:GovRepo.ScriptPath   = $v }
                        'manifestPath' { $script:GovRepo.ManifestPath = $v }
                    }
                }
            }
        }
        $fmt = Get-GovProp $cfg 'format';     if ($fmt) { $script:Gov.Format = $fmt }
        $av  = Get-GovProp $cfg 'apiVersion'; if ($av)  { $script:Gov.ApiVersion = $av }
        $ed  = Get-GovProp $cfg 'exportDir';  if ($ed -and (Test-Path $ed)) { $script:Gov.ExportDir = $ed }
        $sc  = Get-GovProp $cfg 'scopes';     if ($sc)  { $script:Gov.Scopes = @($sc) }
    }
    catch { Write-Host "Could not read config ($p): $($_.Exception.Message)" -ForegroundColor DarkGray }
}

function Save-GovConfig {
    [CmdletBinding()] param()
    $p = Get-GovConfigPath
    $obj = [ordered]@{
        repo = [ordered]@{
            owner        = $script:GovRepo.Owner
            name         = $script:GovRepo.Name
            branch       = $script:GovRepo.Branch
            scriptPath   = $script:GovRepo.ScriptPath
            manifestPath = $script:GovRepo.ManifestPath
        }
        format     = $script:Gov.Format
        apiVersion = $script:Gov.ApiVersion
        exportDir  = $script:Gov.ExportDir
        scopes     = @($script:Gov.Scopes)
    }
    try {
        ($obj | ConvertTo-Json -Depth 6) | Set-Content -Path $p -Encoding UTF8
        Write-Host ("  Saved configuration -> {0}" -f $p) -ForegroundColor Green
    }
    catch { Write-Host "  Save failed: $($_.Exception.Message)" -ForegroundColor Red }
}

function Test-GovRepoConfigured {
    $o = $script:GovRepo.Owner
    $n = $script:GovRepo.Name
    return [bool]($o -and $o.Trim() -and $o -notmatch 'your-org' -and $n -and $n.Trim())
}

function Get-GovRawScriptUrl  { 'https://raw.githubusercontent.com/{0}/{1}/{2}/{3}' -f $script:GovRepo.Owner, $script:GovRepo.Name, $script:GovRepo.Branch, $script:GovRepo.ScriptPath }
function Get-GovManifestUrl   { 'https://raw.githubusercontent.com/{0}/{1}/{2}/{3}' -f $script:GovRepo.Owner, $script:GovRepo.Name, $script:GovRepo.Branch, $script:GovRepo.ManifestPath }
function Get-GovCommitsApiUrl { 'https://api.github.com/repos/{0}/{1}/commits?path={2}&sha={3}&per_page=1' -f $script:GovRepo.Owner, $script:GovRepo.Name, $script:GovRepo.ScriptPath, $script:GovRepo.Branch }
function Get-GovRepoWebUrl    { 'https://github.com/{0}/{1}' -f $script:GovRepo.Owner, $script:GovRepo.Name }

# Hash that ignores line-ending / BOM differences so GitHub (LF) vs local (CRLF)
# does not produce false "update available" results.
function Get-GovNormalizedHash {
    param([string] $Text)
    if ($null -eq $Text) { return $null }
    $norm  = ($Text -replace "`r`n", "`n") -replace "`r", "`n"
    $norm  = $norm.TrimEnd("`n")
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($norm)
    $sha   = [System.Security.Cryptography.SHA256]::Create()
    try { return ((($sha.ComputeHash($bytes)) | ForEach-Object { $_.ToString('x2') }) -join '') }
    finally { $sha.Dispose() }
}

function Get-GovLocalScriptText {
    if ($script:GovScriptPath -and (Test-Path $script:GovScriptPath)) {
        return (Get-Content -Path $script:GovScriptPath -Raw)
    }
    return $null
}

function Get-GovUpdateInfo {
    [CmdletBinding()] param([switch] $Quiet)
    $info = [ordered]@{
        Configured = $false; Checked = (Get-Date); UpdateAvailable = $false; Status = 'Source not configured'
        LocalHash = $null; RemoteHash = $null; RemoteContent = $null
        RemoteSha = $null; RemoteCommitDate = $null; RemoteMessage = $null
        LatestVersion = $null; ReleaseDate = $null; Notes = $null; Error = $null
    }
    if (-not (Test-GovRepoConfigured)) { $script:GovUpdate = [pscustomobject]$info; return $script:GovUpdate }

    $info.Configured = $true
    $info.Status     = 'Checking'
    if (-not $Quiet) { Write-Host '  Checking GitHub for updates...' -ForegroundColor DarkGray }
    try {
        $info.LocalHash = Get-GovNormalizedHash (Get-GovLocalScriptText)
        $resp   = Invoke-WebRequest -Uri (Get-GovRawScriptUrl) -UseBasicParsing -TimeoutSec 20 -Headers @{ 'User-Agent' = 'EntraGovernanceTui' }
        $remote = [string]$resp.Content
        $info.RemoteContent   = $remote
        $info.RemoteHash      = Get-GovNormalizedHash $remote
        $info.UpdateAvailable = [bool]($info.LocalHash -and $info.RemoteHash -and ($info.LocalHash -ne $info.RemoteHash))
        $info.Status          = if ($info.UpdateAvailable) { 'Update available' } else { 'Up to date' }
    }
    catch {
        $info.Error  = $_.Exception.Message
        $info.Status = 'Check failed (offline?)'
        $script:GovUpdate = [pscustomobject]$info
        return $script:GovUpdate
    }
    # Best-effort metadata (commit date/sha + manifest version). Never fatal.
    try {
        $commits = Invoke-RestMethod -Uri (Get-GovCommitsApiUrl) -TimeoutSec 15 -Headers @{ 'User-Agent' = 'EntraGovernanceTui' }
        $c = @($commits)[0]
        if ($c) {
            $info.RemoteSha        = Get-GovProp $c 'sha'
            $commit                = Get-GovProp $c 'commit'
            $committer             = Get-GovProp $commit 'committer'
            $info.RemoteCommitDate = Get-GovProp $committer 'date'
            $msg                   = Get-GovProp $commit 'message'
            if ($msg) { $info.RemoteMessage = ($msg -split "`n")[0] }
        }
    } catch { }
    try {
        $m = Invoke-RestMethod -Uri (Get-GovManifestUrl) -TimeoutSec 10 -Headers @{ 'User-Agent' = 'EntraGovernanceTui' }
        $info.LatestVersion = Get-GovProp $m 'version'
        $info.ReleaseDate   = Get-GovProp $m 'releaseDate'
        $info.Notes         = Get-GovProp $m 'notes'
    } catch { }

    $script:GovUpdate = [pscustomobject]$info
    return $script:GovUpdate
}

function Show-GovVersionBannerLine {
    [CmdletBinding()] param()
    Write-Host ('  Version: {0} ({1})   Graph: {2}   API ref: {3}' -f `
        $script:GovVersion, $script:GovReleaseDate, ($script:GovGraphChannels -join '+'), $script:GovApiVerifiedDate) -ForegroundColor DarkGray
    Write-Host '  Update : ' -NoNewline
    $u = $script:GovUpdate
    if (-not $u)             { Write-Host "press 'U' to check GitHub (or 'V' for About)" -ForegroundColor DarkGray; return }
    if (-not $u.Configured)  { Write-Host "self-update source not set - configure in About ('V')" -ForegroundColor DarkGray; return }
    switch ($u.Status) {
        'Update available' {
            $ver = if ($u.LatestVersion) { " (v$($u.LatestVersion))" } else { '' }
            Write-Host ("UPDATE AVAILABLE{0} - press 'U' to download & apply" -f $ver) -ForegroundColor Yellow
        }
        'Up to date' { Write-Host ("up to date (checked {0:HH:mm})" -f $u.Checked) -ForegroundColor Green }
        default      { Write-Host $u.Status -ForegroundColor DarkGray }
    }
}

function Invoke-GovSelfUpdate {
    [CmdletBinding()] param([switch] $Force)
    if (-not (Test-GovRepoConfigured)) {
        Write-Host "  Self-update source not configured. Open About ('V') -> Set repo." -ForegroundColor Yellow
        return
    }
    $u = $script:GovUpdate
    if (-not $u -or -not $u.RemoteContent) { $u = Get-GovUpdateInfo }
    if (-not $u.RemoteContent) { Write-Host "  Could not retrieve latest script: $($u.Error)" -ForegroundColor Red; return }
    if (-not $u.UpdateAvailable -and -not $Force) { Write-Host '  Already up to date.' -ForegroundColor Green; return }

    # Validate the download before overwriting anything.
    $perrors = $null; $ptokens = $null
    try { [void][System.Management.Automation.Language.Parser]::ParseInput($u.RemoteContent, [ref]$ptokens, [ref]$perrors) } catch { }
    if ($perrors -and @($perrors).Count -gt 0) {
        Write-Host '  Downloaded script has parse errors; aborting to protect your install.' -ForegroundColor Red; return
    }
    if ($u.RemoteContent.Length -lt 1000 -or $u.RemoteContent -notmatch 'Start-GovTui') {
        Write-Host "  Downloaded content does not look like this tool; aborting." -ForegroundColor Red; return
    }
    if (-not $script:GovScriptPath -or -not (Test-Path $script:GovScriptPath)) {
        Write-Host '  Cannot locate the local script file to replace.' -ForegroundColor Red; return
    }

    $backup = "$script:GovScriptPath.bak-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    try {
        Copy-Item -Path $script:GovScriptPath -Destination $backup -Force
        [System.IO.File]::WriteAllText($script:GovScriptPath, $u.RemoteContent, [System.Text.UTF8Encoding]::new($false))
        Write-Host ''
        Write-Host '  Update applied successfully.' -ForegroundColor Green
        Write-Host ("  Backup saved : {0}" -f $backup) -ForegroundColor DarkGray
        Write-Host '  Restart the tool to load the new version.' -ForegroundColor Yellow
        [void](Get-GovUpdateInfo -Quiet)
        $ans = Read-Host '  Relaunch now in a new window? (y/N)'
        if ($ans -match '^(y|yes)$') {
            try {
                Start-Process pwsh -ArgumentList @('-NoExit', '-File', $script:GovScriptPath)
                Write-Host '  Launched a new instance - this one will exit.' -ForegroundColor Cyan
                $script:GovRequestExit = $true
            }
            catch { Write-Host "  Could not relaunch automatically: $($_.Exception.Message)" -ForegroundColor Yellow }
        }
    }
    catch {
        Write-Host "  Update failed: $($_.Exception.Message)" -ForegroundColor Red
        if (Test-Path $backup) { Write-Host "  Previous version intact (backup: $backup)." -ForegroundColor DarkGray }
    }
}

function Set-GovRepoInteractive {
    [CmdletBinding()] param()
    Write-Host ''
    $o = Read-Host ("   GitHub owner/org      [{0}]" -f $script:GovRepo.Owner)
    if ($o) { $script:GovRepo.Owner = $o.Trim() }
    $n = Read-Host ("   Repository name       [{0}]" -f $script:GovRepo.Name)
    if ($n) { $script:GovRepo.Name = $n.Trim() }
    $b = Read-Host ("   Branch                [{0}]" -f $script:GovRepo.Branch)
    if ($b) { $script:GovRepo.Branch = $b.Trim() }
    $s = Read-Host ("   Script path in repo   [{0}]" -f $script:GovRepo.ScriptPath)
    if ($s) { $script:GovRepo.ScriptPath = $s.Trim() }
    Write-Host '   Repo updated. (W) saves it to config; (C) checks now.' -ForegroundColor Green
    [void](Get-GovUpdateInfo -Quiet)
    Read-Host '   Press Enter to continue' | Out-Null
}

function Show-GovAboutMenu {
    [CmdletBinding()] param()
    while ($true) {
        Clear-GovHost
        Show-GovBanner
        Write-Host '  ABOUT / VERSION / UPDATE' -ForegroundColor White
        Write-GovRule
        Write-Host ("   Tool version      : {0}" -f $script:GovVersion)
        Write-Host ("   Released          : {0}" -f $script:GovReleaseDate)
        Write-Host ("   Graph channels    : {0}" -f ($script:GovGraphChannels -join ', '))
        Write-Host ("   API ref verified  : {0}  (endpoints + scopes vs learn.microsoft.com)" -f $script:GovApiVerifiedDate)
        Write-Host ("   Script path       : {0}" -f $script:GovScriptPath)
        Write-Host ("   Config file       : {0}" -f (Get-GovConfigPath))
        Write-GovRule
        Write-Host '   GitHub self-update source' -ForegroundColor White
        if (Test-GovRepoConfigured) {
            Write-Host ("   Repo              : {0}  (branch {1})" -f (Get-GovRepoWebUrl), $script:GovRepo.Branch)
            Write-Host ("   Script in repo    : {0}" -f $script:GovRepo.ScriptPath)
        }
        else { Write-Host '   Repo              : <not configured>' -ForegroundColor Yellow }
        $u = $script:GovUpdate
        if ($u -and $u.Configured) {
            $col = if ($u.UpdateAvailable) { 'Yellow' } else { 'Green' }
            Write-Host ("   Last checked      : {0}" -f $u.Checked)
            Write-Host ("   Status            : {0}" -f $u.Status) -ForegroundColor $col
            if ($u.RemoteCommitDate) {
                $shortSha = if ($u.RemoteSha) { $u.RemoteSha.Substring(0, [Math]::Min(7, $u.RemoteSha.Length)) } else { '' }
                Write-Host ("   Remote commit     : {0}  {1}" -f $u.RemoteCommitDate, $shortSha)
            }
            if ($u.RemoteMessage) { Write-Host ("   Remote message    : {0}" -f $u.RemoteMessage) -ForegroundColor DarkGray }
            if ($u.LatestVersion) { Write-Host ("   Manifest version  : {0} ({1})" -f $u.LatestVersion, $u.ReleaseDate) }
        }
        Write-GovRule
        Write-Host ("   Note: {0}" -f $script:GovApiNotes) -ForegroundColor DarkGray
        Write-GovRule
        Write-Host '   C. Check for updates now      U. Download & apply update'
        Write-Host '   S. Set GitHub repo            W. Save config        B. Back'
        Write-GovRule
        $sel = (Read-Host '   about>').Trim().ToLower()
        switch ($sel) {
            'c' { [void](Get-GovUpdateInfo); Read-Host '   Press Enter to continue' | Out-Null }
            'u' { Invoke-GovSelfUpdate; if (-not $script:GovRequestExit) { Read-Host '   Press Enter to continue' | Out-Null } }
            's' { Set-GovRepoInteractive }
            'w' { Save-GovConfig; Read-Host '   Press Enter to continue' | Out-Null }
            'b' { return }
            ''  { }
            default { }
        }
        if ($script:GovRequestExit) { return }
    }
}

#endregion

#region ----------------------------------------------------------- UI helpers

$script:GovColors = @{
    Title='Cyan'; Ok='Green'; Warn='Yellow'; Err='Red'; Num='Yellow'; Dim='DarkGray'; Head='White'
}

function Write-GovRule {
    param([string]$Char = '-', [int]$Width = 78, [string]$Color = 'DarkGray')
    Write-Host ($Char * $Width) -ForegroundColor $Color
}

# Resilient screen clear: never throws (e.g. when stdout is redirected / no console).
function Clear-GovHost {
    try { Clear-Host }
    catch {
        try { [Console]::Clear() } catch { Write-Host "`n" }
    }
}

function Write-GovCenter {
    param([string]$Text, [int]$Width = 78, [string]$Color = 'Cyan')
    if ($Text.Length -ge $Width) { Write-Host $Text -ForegroundColor $Color; return }
    $pad = [int](($Width - $Text.Length) / 2)
    Write-Host ((' ' * $pad) + $Text) -ForegroundColor $Color
}

function Get-GovContext {
    [CmdletBinding()] param()
    try { return Get-MgContext } catch { return $null }
}function Show-GovBanner {
    [CmdletBinding()] param()
    $ctx = Get-GovContext
    Write-GovRule '=' 78 'Cyan'
    Write-GovCenter 'ENTRA  TENANT  GOVERNANCE  —  TUI' 78 'Cyan'
    Write-GovCenter 'Identity Governance management via Microsoft Graph' 78 'DarkGray'
    Write-GovRule '=' 78 'Cyan'

    if ($ctx -and $ctx.AuthType) {
        $who = if ($ctx.Account) { $ctx.Account }
               elseif ($ctx.AppName) { "$($ctx.AppName) (app)" }
               elseif ($ctx.ClientId) { "$($ctx.ClientId) (app)" }
               else { '<unknown>' }
        $scopeCount = if ($ctx.Scopes) { @($ctx.Scopes).Count } else { 0 }
        Write-Host '  Status : ' -NoNewline; Write-Host 'CONNECTED' -ForegroundColor Green
        Write-Host '  Account: ' -NoNewline; Write-Host $who -ForegroundColor White
        Write-Host '  Tenant : ' -NoNewline; Write-Host ($ctx.TenantId  ?? '<n/a>') -ForegroundColor White
        Write-Host '  Auth   : ' -NoNewline; Write-Host ("{0}   Scopes: {1}   Graph: {2}" -f $ctx.AuthType, $scopeCount, (Get-GovEffectiveVersion)) -ForegroundColor White
    }
    else {
        Write-Host '  Status : ' -NoNewline; Write-Host 'NOT CONNECTED' -ForegroundColor Red
        Write-Host '  Action : ' -NoNewline; Write-Host "Choose 'L' (Login) to sign in. Interactive browser sign-in is the default." -ForegroundColor Yellow
    }
    Show-GovPrereqBannerLine
    Show-GovVersionBannerLine
    Write-GovRule '-' 78 'DarkGray'
}

function Get-GovEffectiveVersion {
    param([object]$Action)
    switch ($script:Gov.ApiVersion) {
        'v1.0' { return 'v1.0' }
        'beta' { return 'beta' }
        default {
            if ($Action -and $Action.Version) { return $Action.Version }
            return 'v1.0'
        }
    }
}

function Get-GovBreadcrumb {
    param([System.Collections.Generic.Stack[object]]$Stack)
    $titles = @($Stack.ToArray() | ForEach-Object { $_.Title })
    [array]::Reverse($titles)
    return ($titles -join ' > ')
}

#endregion

#region --------------------------------------------------------- Module / auth

function Confirm-GovModule {
    [CmdletBinding()] param([switch] $Quiet)
    $m = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Select-Object -First 1
    if (-not $m) {
        if (-not $Quiet) {
            Write-Host 'Microsoft.Graph.Authentication is not installed.' -ForegroundColor Red
            Write-Host "Open the Prerequisites menu (press 'P') to install it automatically." -ForegroundColor Yellow
        }
        return $false
    }
    try { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop; return $true }
    catch {
        if (-not $Quiet) { Write-Host "Failed to import Microsoft.Graph.Authentication: $($_.Exception.Message)" -ForegroundColor Red }
        return $false
    }
}

function Connect-Gov {
    [CmdletBinding()]
    param(
        [ValidateSet('Interactive','DeviceCode','AppSecret','AppCertificate','ManagedIdentity')]
        [string]   $Mode = 'Interactive',
        [string]   $TenantId,
        [string[]] $Scopes,
        [string]   $ClientId,
        [string]   $CertThumbprint
    )
    if (-not (Confirm-GovModule)) {
        Write-Host "Microsoft.Graph.Authentication is required. Open Prerequisites (press 'P') to install it." -ForegroundColor Red
        return
    }
    $scopeList = if ($Scopes) { $Scopes } else { $script:Gov.Scopes }

    try {
        switch ($Mode) {
            'Interactive'  {
                $p = @{ Scopes = $scopeList; NoWelcome = $true }
                if ($TenantId) { $p.TenantId = $TenantId }
                Connect-MgGraph @p
            }
            'DeviceCode'   {
                $p = @{ Scopes = $scopeList; UseDeviceCode = $true; NoWelcome = $true }
                if ($TenantId) { $p.TenantId = $TenantId }
                Connect-MgGraph @p
            }
            'AppCertificate' {
                if (-not $ClientId)       { $ClientId       = Read-Host 'Application (client) Id' }
                if (-not $TenantId)       { $TenantId       = Read-Host 'Tenant Id' }
                if (-not $CertThumbprint) { $CertThumbprint = Read-Host 'Certificate thumbprint' }
                Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertThumbprint -NoWelcome
            }
            'AppSecret'    {
                if (-not $ClientId) { $ClientId = Read-Host 'Application (client) Id' }
                if (-not $TenantId) { $TenantId = Read-Host 'Tenant Id' }
                $sec  = Read-Host 'Client secret' -AsSecureString
                $cred = [pscredential]::new($ClientId, $sec)
                Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $cred -NoWelcome
            }
            'ManagedIdentity' {
                if ($ClientId) { Connect-MgGraph -Identity -ClientId $ClientId -NoWelcome }
                else           { Connect-MgGraph -Identity -NoWelcome }
            }
        }
        $ctx = Get-GovContext
        if ($ctx) {
            $script:Gov.Scopes = $scopeList
            Write-Host ("Signed in: {0}" -f ($ctx.Account ?? $ctx.AppName ?? $ctx.ClientId)) -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Sign-in failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Disconnect-Gov {
    [CmdletBinding()] param()
    try { Disconnect-MgGraph -ErrorAction Stop | Out-Null; Write-Host 'Signed out.' -ForegroundColor Yellow }
    catch { Write-Host 'No active session to sign out.' -ForegroundColor DarkGray }
}

function Show-GovLoginMenu {
    [CmdletBinding()] param()
    Clear-GovHost
    Show-GovBanner
    Write-Host '  LOGIN / ACCOUNT' -ForegroundColor White
    Write-GovRule
    $opts = @(
        '1. Interactive sign-in (browser)        [default, dynamic]'
        '2. Device code sign-in                  [headless / remote]'
        '3. App-only with certificate            [ClientId + Tenant + Thumbprint]'
        '4. App-only with client secret          [ClientId + Tenant + Secret]'
        '5. Managed identity                     [Azure-hosted]'
        '6. Sign out'
        '7. Show full context (who am I)'
        '8. Set sign-in scopes'
    )
    foreach ($o in $opts) {
        Write-Host '   ' -NoNewline
        Write-Host $o.Substring(0,2) -ForegroundColor Yellow -NoNewline
        Write-Host $o.Substring(2)
    }
    Write-GovRule
    $sel = Read-Host '  Select 1-8 (or b to go back)'
    switch ($sel.Trim()) {
        '1' { Connect-Gov -Mode Interactive  -TenantId $script:GovTenant }
        '2' { Connect-Gov -Mode DeviceCode   -TenantId $script:GovTenant }
        '3' { Connect-Gov -Mode AppCertificate }
        '4' { Connect-Gov -Mode AppSecret }
        '5' { Connect-Gov -Mode ManagedIdentity }
        '6' { Disconnect-Gov }
        '7' { Show-GovWhoAmI }
        '8' {
            $raw = Read-Host '  Space/comma separated scopes (blank = default set)'
            if ([string]::IsNullOrWhiteSpace($raw)) { $script:Gov.Scopes = $script:Gov.DefaultScopes }
            else { $script:Gov.Scopes = @($raw -split '[,\s]+' | Where-Object { $_ }) }
            Write-Host ("Scopes set ({0})." -f @($script:Gov.Scopes).Count) -ForegroundColor Green
        }
        default { return }
    }
    Read-Host '  Press Enter to continue' | Out-Null
}

function Show-GovWhoAmI {
    [CmdletBinding()] param()
    $ctx = Get-GovContext
    if (-not $ctx) { Write-Host 'Not connected.' -ForegroundColor Red; return }
    Write-Host ''
    Write-Host 'Current Graph context:' -ForegroundColor White
    $ctx | Select-Object Account, AppName, ClientId, TenantId, AuthType, TokenCredentialType, ContextScope | Format-List | Out-Host
    Write-Host 'Granted scopes:' -ForegroundColor White
    @($ctx.Scopes) | Sort-Object | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
}

#endregion

#region ----------------------------------------------------------- Data access

function Get-GovActionUri {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Action)

    $ver  = Get-GovEffectiveVersion -Action $Action
    $base = "$($script:Gov.GraphBaseUrl)/$ver/$($Action.Path.TrimStart('/'))"

    $q = @()
    if ($script:Gov.Top -gt 0)              { $q += ('$top='     + $script:Gov.Top) }
    $filter  = $script:Gov.Filter  ?? $Action.Filter
    $select  = $script:Gov.Select  ?? $Action.Select
    $expand  = $script:Gov.Expand  ?? $Action.Expand
    $orderby = $script:Gov.OrderBy ?? $Action.OrderBy
    if ($filter)  { $q += ('$filter='  + ($filter  -replace ' ', '%20')) }
    if ($select)  { $q += ('$select='  + $select) }
    if ($expand)  { $q += ('$expand='  + $expand) }
    if ($orderby) { $q += ('$orderby=' + ($orderby -replace ' ', '%20')) }

    if ($q.Count) { return ($base + '?' + ($q -join '&')) }
    return $base
}

function Invoke-GovGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $Uri,
        [int] $MaxItems = 200,
        [switch] $Advanced   # adds ConsistencyLevel: eventual for $count/$search
    )
    $headers = @{}
    if ($Advanced) { $headers['ConsistencyLevel'] = 'eventual' }

    $items = [System.Collections.Generic.List[object]]::new()
    $next  = $Uri
    $pages = 0
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject -Headers $headers -ErrorAction Stop
        $pages++
        $propNames = @($resp.PSObject.Properties.Name)
        if ($propNames -contains 'value') {
            foreach ($it in @($resp.value)) {
                $items.Add($it)
                if ($items.Count -ge $MaxItems) { break }
            }
            $next = if ($items.Count -ge $MaxItems) { $null } else { $resp.PSObject.Properties['@odata.nextLink'].Value }
        }
        else {
            # single resource (e.g. settings, organization element)
            $items.Add($resp)
            $next = $null
        }
        if ($pages -gt 1000) { break }   # safety valve
    }
    return ,($items.ToArray())
}

#endregion

#region ------------------------------------------------------ Output / export

function Get-GovDisplayColumns {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Item, [int]$Max = 8)
    $cols = [System.Collections.Generic.List[string]]::new()
    foreach ($p in $Item.PSObject.Properties) {
        if ($p.Name -like '@odata*') { continue }
        $v = $p.Value
        $isScalar = ($null -eq $v) -or
                    ($v -is [string]) -or ($v -is [bool]) -or ($v -is [datetime]) -or
                    ($v -is [int]) -or ($v -is [long]) -or ($v -is [double]) -or ($v -is [decimal])
        if ($isScalar) { $cols.Add($p.Name) }
    }
    if ($cols.Count -eq 0) {
        foreach ($p in $Item.PSObject.Properties) { if ($p.Name -notlike '@odata*') { $cols.Add($p.Name) } }
    }
    return @($cols | Select-Object -First $Max)
}

function ConvertTo-GovFlat {
    [CmdletBinding()]
    param([Parameter(Mandatory, ValueFromPipeline)] $Item)
    process {
        $o = [ordered]@{}
        foreach ($p in $Item.PSObject.Properties) {
            if ($p.Name -like '@odata*') { continue }
            $v = $p.Value
            if ($null -eq $v) { $o[$p.Name] = $null }
            elseif ($v -is [string] -or $v -is [bool] -or $v -is [datetime] -or $v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal]) {
                $o[$p.Name] = $v
            }
            else { $o[$p.Name] = ($v | ConvertTo-Json -Depth 8 -Compress) }
        }
        [pscustomobject]$o
    }
}

function Show-GovResult {
    [CmdletBinding()]
    param(
        [object[]] $Data,
        [string]   $Format,
        [string[]] $Columns
    )
    if (-not $Format) { $Format = $script:Gov.Format }
    if ($null -eq $Data -or @($Data).Count -eq 0) {
        Write-Host 'No items returned.' -ForegroundColor Yellow
        return
    }
    $Data = @($Data)
    switch ($Format) {
        'Table' {
            $cols = if ($Columns) { $Columns } elseif ($script:Gov.LastAction -and $script:Gov.LastAction.Columns) { $script:Gov.LastAction.Columns } else { Get-GovDisplayColumns $Data[0] }
            $cols = @($cols | Where-Object { $_ })
            $Data | Select-Object $cols | Format-Table -AutoSize -Wrap | Out-Host
        }
        'List' { $Data | Format-List | Out-Host }
        'Json' { ($Data | ConvertTo-Json -Depth 12) | Out-Host }
        'Csv'  { ($Data | ConvertTo-GovFlat | ConvertTo-Csv -NoTypeInformation) | Out-Host }
        'Grid' {
            $ocgv = Get-Command Out-ConsoleGridView -ErrorAction SilentlyContinue
            $ogv  = Get-Command Out-GridView        -ErrorAction SilentlyContinue
            if     ($ocgv) { $Data | ConvertTo-GovFlat | Out-ConsoleGridView -Title 'Governance result' }
            elseif ($ogv)  { $Data | ConvertTo-GovFlat | Out-GridView        -Title 'Governance result' }
            else { Write-Host 'Grid view unavailable; falling back to Table.' -ForegroundColor Yellow; Show-GovResult -Data $Data -Format Table }
        }
        default { $Data | Format-Table -AutoSize | Out-Host }
    }
    Write-Host ("({0} item(s) shown, format={1})" -f @($Data).Count, $Format) -ForegroundColor DarkGray
}

function Export-GovResult {
    [CmdletBinding()]
    param([string] $Path, [string] $Format)

    $data = @($script:Gov.LastView)
    if ($data.Count -eq 0) { Write-Host 'Nothing to export — run an action first.' -ForegroundColor Yellow; return }

    if ($Path -and $Path.Trim().ToLower() -eq 'clip') {
        ($data | ConvertTo-Json -Depth 12) | Set-Clipboard
        Write-Host ("Copied {0} item(s) as JSON to clipboard." -f $data.Count) -ForegroundColor Green
        return
    }

    if (-not $Path) {
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $name  = if ($script:Gov.LastAction) { ($script:Gov.LastAction.Title -replace '[^\w]+','_') } else { 'governance' }
        $ext   = switch (($Format ?? $script:Gov.Format)) { 'Json' {'json'} 'Csv' {'csv'} 'Table' {'csv'} default {'json'} }
        $Path  = Join-Path $script:Gov.ExportDir ("{0}_{1}.{2}" -f $name, $stamp, $ext)
    }
    if (-not [System.IO.Path]::IsPathRooted($Path)) { $Path = Join-Path $script:Gov.ExportDir $Path }

    $ext = [System.IO.Path]::GetExtension($Path).TrimStart('.').ToLower()
    try {
        switch ($ext) {
            'json' { ($data | ConvertTo-Json -Depth 12) | Set-Content -Path $Path -Encoding UTF8 }
            'csv'  { $data | ConvertTo-GovFlat | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 }
            'txt'  { ($data | Format-List | Out-String) | Set-Content -Path $Path -Encoding UTF8 }
            default { ($data | ConvertTo-Json -Depth 12) | Set-Content -Path $Path -Encoding UTF8 }
        }
        Write-Host ("Exported {0} item(s) -> {1}" -f $data.Count, $Path) -ForegroundColor Green
    }
    catch { Write-Host "Export failed: $($_.Exception.Message)" -ForegroundColor Red }
}

#endregion

#region ------------------------------------------------------------- Menu tree

function New-GovAction {
    param(
        [Parameter(Mandatory)][string] $Title,
        [Parameter(Mandatory)][string] $Path,
        [string]   $Version = 'v1.0',
        [string[]] $Columns,
        [string]   $Select,
        [string]   $Expand,
        [string]   $OrderBy,
        [string]   $Filter,
        [string]   $Scope,
        [switch]   $RequiresFilter,
        [switch]   $Singleton,
        [string]   $Description
    )
    [ordered]@{
        Type='action'; Title=$Title; Path=$Path; Version=$Version
        Columns=$Columns; Select=$Select; Expand=$Expand; OrderBy=$OrderBy; Filter=$Filter
        Scope=$Scope; RequiresFilter=[bool]$RequiresFilter; Singleton=[bool]$Singleton
        Description=$Description
    }
}

function New-GovMenu {
    param(
        [Parameter(Mandatory)][string] $Title,
        [Parameter(Mandatory)][object[]] $Items,
        [string] $Description
    )
    [ordered]@{ Type='menu'; Title=$Title; Items=$Items; Description=$Description }
}

function Get-GovMenuTree {
    [CmdletBinding()] param()

    $entitlement = New-GovMenu -Title 'Entitlement Management' -Items @(
        New-GovAction 'Access packages'          'identityGovernance/entitlementManagement/accessPackages'        -Scope 'EntitlementManagement.Read.All' -Columns id,displayName,isHidden,createdDateTime,modifiedDateTime
        New-GovAction 'Access package catalogs'  'identityGovernance/entitlementManagement/catalogs'              -Scope 'EntitlementManagement.Read.All' -Columns id,displayName,catalogType,state,isExternallyVisible
        New-GovAction 'Assignments'              'identityGovernance/entitlementManagement/assignments'           -Scope 'EntitlementManagement.Read.All' -Columns id,state,status,createdDateTime -Expand 'accessPackage($select=displayName),target'
        New-GovAction 'Assignment requests'      'identityGovernance/entitlementManagement/assignmentRequests'    -Scope 'EntitlementManagement.Read.All' -Columns id,requestType,state,status,createdDateTime
        New-GovAction 'Assignment policies'      'identityGovernance/entitlementManagement/assignmentPolicies'    -Scope 'EntitlementManagement.Read.All' -Columns id,displayName,allowedTargetScope,createdDateTime
        New-GovAction 'Connected organizations'  'identityGovernance/entitlementManagement/connectedOrganizations' -Scope 'EntitlementManagement.Read.All' -Columns id,displayName,state,createdDateTime
        New-GovAction 'Resource requests'        'identityGovernance/entitlementManagement/resourceRequests'      -Scope 'EntitlementManagement.Read.All' -Columns id,requestType,requestState
        New-GovAction 'Settings (singleton)'     'identityGovernance/entitlementManagement/settings'              -Scope 'EntitlementManagement.Read.All' -Singleton
    )

    $reviews = New-GovMenu -Title 'Access Reviews' -Items @(
        New-GovAction 'Review definitions'       'identityGovernance/accessReviews/definitions'        -Scope 'AccessReview.Read.All'      -Columns id,displayName,status,createdDateTime
        New-GovAction 'History definitions (RW scope, last 30d)' 'identityGovernance/accessReviews/historyDefinitions' -Scope 'AccessReview.ReadWrite.All' -Columns id,displayName,status,createdDateTime
    )

    $pimDir = New-GovMenu -Title 'PIM - Directory Roles' -Items @(
        New-GovAction 'Role definitions'                 'roleManagement/directory/roleDefinitions'                  -Scope 'RoleManagement.Read.All' -Columns id,displayName,isBuiltIn,isEnabled
        New-GovAction 'Active role assignments'          'roleManagement/directory/roleAssignments'                  -Scope 'RoleManagement.Read.All' -Columns id,principalId,roleDefinitionId,directoryScopeId
        New-GovAction 'Assignment schedules'             'roleManagement/directory/roleAssignmentSchedules'          -Scope 'RoleManagement.Read.All' -Columns id,principalId,roleDefinitionId,assignmentType,status
        New-GovAction 'Assignment schedule instances'    'roleManagement/directory/roleAssignmentScheduleInstances'  -Scope 'RoleManagement.Read.All' -Columns id,principalId,roleDefinitionId,assignmentType,startDateTime,endDateTime
        New-GovAction 'Assignment schedule requests'     'roleManagement/directory/roleAssignmentScheduleRequests'   -Scope 'RoleAssignmentSchedule.ReadWrite.Directory' -Columns id,action,principalId,roleDefinitionId,status,createdDateTime
        New-GovAction 'Eligibility schedules'            'roleManagement/directory/roleEligibilitySchedules'         -Scope 'RoleManagement.Read.All' -Columns id,principalId,roleDefinitionId,memberType,status
        New-GovAction 'Eligibility schedule instances'   'roleManagement/directory/roleEligibilityScheduleInstances' -Scope 'RoleManagement.Read.All' -Columns id,principalId,roleDefinitionId,memberType,startDateTime,endDateTime
        New-GovAction 'Eligibility schedule requests'    'roleManagement/directory/roleEligibilityScheduleRequests'  -Scope 'RoleEligibilitySchedule.ReadWrite.Directory' -Columns id,action,principalId,roleDefinitionId,status,createdDateTime
    )

    # PIM for Groups list endpoints REQUIRE a $filter on groupId or principalId.
    $pimGrp = New-GovMenu -Title 'PIM for Groups' -Items @(
        New-GovAction 'Eligibility schedules'            'identityGovernance/privilegedAccess/group/eligibilitySchedules'         -Scope 'PrivilegedEligibilitySchedule.Read.AzureADGroup' -RequiresFilter -Columns id,principalId,groupId,accessId,memberType,status
        New-GovAction 'Eligibility schedule instances'   'identityGovernance/privilegedAccess/group/eligibilityScheduleInstances' -Scope 'PrivilegedEligibilitySchedule.Read.AzureADGroup' -RequiresFilter -Columns id,principalId,groupId,accessId,startDateTime,endDateTime
        New-GovAction 'Eligibility schedule requests'    'identityGovernance/privilegedAccess/group/eligibilityScheduleRequests'  -Scope 'PrivilegedEligibilitySchedule.Read.AzureADGroup' -RequiresFilter -Columns id,principalId,groupId,accessId,status
        New-GovAction 'Assignment schedules'             'identityGovernance/privilegedAccess/group/assignmentSchedules'          -Scope 'PrivilegedAssignmentSchedule.Read.AzureADGroup' -RequiresFilter -Columns id,principalId,groupId,accessId,assignmentType,status
        New-GovAction 'Assignment schedule instances'    'identityGovernance/privilegedAccess/group/assignmentScheduleInstances'  -Scope 'PrivilegedAssignmentSchedule.Read.AzureADGroup' -RequiresFilter -Columns id,principalId,groupId,accessId,startDateTime,endDateTime
        New-GovAction 'Assignment schedule requests'     'identityGovernance/privilegedAccess/group/assignmentScheduleRequests'   -Scope 'PrivilegedAssignmentSchedule.Read.AzureADGroup' -RequiresFilter -Columns id,principalId,groupId,accessId,status
    )

    $lifecycle = New-GovMenu -Title 'Lifecycle Workflows' -Items @(
        New-GovAction 'Workflows'                'identityGovernance/lifecycleWorkflows/workflows'             -Scope 'LifecycleWorkflows.Read.All' -Columns id,displayName,category,isEnabled,isSchedulingEnabled,createdDateTime
        New-GovAction 'Deleted workflows'        'identityGovernance/lifecycleWorkflows/deletedItems/workflows' -Scope 'LifecycleWorkflows.Read.All' -Columns id,displayName,category,deletedDateTime
        New-GovAction 'Workflow templates'       'identityGovernance/lifecycleWorkflows/workflowTemplates'     -Scope 'LifecycleWorkflows.Read.All' -Columns id,displayName,category
        New-GovAction 'Task definitions'         'identityGovernance/lifecycleWorkflows/taskDefinitions'       -Scope 'LifecycleWorkflows.Read.All' -Columns id,displayName,category,continueOnError,version
        New-GovAction 'Custom task extensions'   'identityGovernance/lifecycleWorkflows/customTaskExtensions'  -Scope 'LifecycleWorkflows.Read.All' -Columns id,displayName,createdDateTime
        New-GovAction 'Settings (singleton)'     'identityGovernance/lifecycleWorkflows/settings'              -Scope 'LifecycleWorkflows.Read.All' -Singleton
    )

    $tou = New-GovMenu -Title 'Terms of Use' -Items @(
        New-GovAction 'Agreements'               'identityGovernance/termsOfUse/agreements' -Scope 'Agreement.Read.All' -Columns id,displayName,isViewingBeforeAcceptanceRequired,isPerDeviceAcceptanceRequired
    )

    $consent = New-GovMenu -Title 'App Consent' -Items @(
        New-GovAction 'App consent requests'     'identityGovernance/appConsent/appConsentRequests' -Scope 'ConsentRequest.Read.All' -Columns id,appId,appDisplayName,consentType
    )

    $tenant = New-GovMenu -Title 'Tenant Overview' -Items @(
        New-GovAction 'Organization'             'organization'    -Scope 'Directory.Read.All' -Columns id,displayName,city,countryLetterCode,tenantType,createdDateTime
        New-GovAction 'Subscribed SKUs'          'subscribedSkus'  -Scope 'Directory.Read.All' -Columns skuId,skuPartNumber,capabilityStatus,consumedUnits
        New-GovAction 'Directory roles (active)' 'directoryRoles'  -Scope 'Directory.Read.All' -Columns id,displayName,roleTemplateId
    )

    return New-GovMenu -Title 'Governance' -Items @(
        $entitlement, $reviews, $pimDir, $pimGrp, $lifecycle, $tou, $consent, $tenant
    )
}

#endregion

#region ----------------------------------------------------------- Action run

function Invoke-GovAction {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Action)

    if (-not (Get-GovContext)) {
        Write-Host ''
        Write-Host 'You are not signed in. Choose "L" from the main menu to log in first.' -ForegroundColor Red
        Read-Host 'Press Enter to continue' | Out-Null
        return
    }

    # Start each action with a clean query slate so overrides do not leak across actions.
    $script:Gov.Filter = $null; $script:Gov.Select = $null; $script:Gov.Expand = $null; $script:Gov.OrderBy = $null
    $script:Gov.LastAction = $Action

    # PIM for Groups list endpoints REQUIRE a $filter on groupId or principalId.
    if ($Action.RequiresFilter) {
        if (-not (Set-GovRequiredFilter -Action $Action)) { return }
    }

    $uri = Get-GovActionUri -Action $Action

    try {
        Write-Host ''
        Write-Host ("GET {0}" -f $uri) -ForegroundColor DarkGray
        $data = Invoke-GovGraph -Uri $uri -MaxItems $script:Gov.MaxItems -Advanced:([bool]($script:Gov.Filter -or $script:Gov.OrderBy))
        $script:Gov.LastResult = @($data)
        $script:Gov.LastView   = @($data)
    }
    catch {
        $msg = $_.Exception.Message
        Write-Host ''
        Write-Host "Request failed: $msg" -ForegroundColor Red
        if ($Action.Scope) { Write-Host ("Required delegated scope for this endpoint: {0}" -f $Action.Scope) -ForegroundColor Yellow }
        if ($msg -match '403|Authorization|Forbidden|privilege|Accepted') {
            Write-Host "Likely a missing scope or directory role. Use Login (L) to (re)consent, or add the scope above via 'Set sign-in scopes'." -ForegroundColor Yellow
        }
        elseif ($msg -match '400|filter') {
            Write-Host "This endpoint may require a valid \$filter. Use the 'filter' command (e.g. filter principalId eq '<guid>')." -ForegroundColor Yellow
        }
        Read-Host 'Press Enter to continue' | Out-Null
        return
    }

    Invoke-GovResultLoop -Action $Action
}

# Prompts for the mandatory PIM-for-Groups filter and stores it as the run filter.
# Returns $true if a filter is set, $false if the user cancelled.
function Set-GovRequiredFilter {
    [CmdletBinding()] param([Parameter(Mandatory)] $Action)
    Clear-GovHost
    Show-GovBanner
    Write-Host ("  {0}" -f $Action.Title) -ForegroundColor White
    Write-GovRule
    Write-Host '  PIM for Groups list APIs require a $filter on groupId or principalId.' -ForegroundColor Yellow
    Write-Host '  Enter one of:' -ForegroundColor DarkGray
    Write-Host "     <group-guid>             -> filters groupId eq '<guid>'" -ForegroundColor DarkGray
    Write-Host "     p <principal-guid>       -> filters principalId eq '<guid>'" -ForegroundColor DarkGray
    Write-Host "     <raw OData filter>       -> used as-is" -ForegroundColor DarkGray
    Write-GovRule
    $val = (Read-Host '  filter (b to cancel)').Trim()
    if (-not $val -or $val.ToLower() -eq 'b') { return $false }

    $guid = '^[0-9a-fA-F-]{36}$'
    if ($val -match $guid) {
        $script:Gov.Filter = "groupId eq '$val'"
    }
    elseif ($val -match '^(p|principal)\s+(.+)$') {
        $script:Gov.Filter = "principalId eq '$($Matches[2].Trim())'"
    }
    else {
        $script:Gov.Filter = $val
    }
    Write-Host ("  Using filter: {0}" -f $script:Gov.Filter) -ForegroundColor Green
    return $true
}

function Invoke-GovResultLoop {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Action)

    while ($true) {
        Clear-GovHost
        Show-GovBanner
        Write-Host ("  RESULT: {0}" -f $Action.Title) -ForegroundColor White
        $opts = "format <t|l|json|csv|grid> | cols a,b | top N | filter <odata> | select | expand | orderby | find <txt> | open N | refresh | export [path|clip] | b"
        Write-Host ("  Commands: {0}" -f $opts) -ForegroundColor DarkGray
        Write-GovRule
        Show-GovResult -Data $script:Gov.LastView
        Write-GovRule

        $in = Read-Host ("  [{0}|max{1}|{2}] result>" -f $script:Gov.Format, $script:Gov.MaxItems, (Get-GovEffectiveVersion -Action $Action))
        $in = $in.Trim()
        if ($in -eq '') { continue }
        $verb = ($in -split '\s+', 2)[0].ToLower()
        $arg  = if ($in -match '\s') { ($in -split '\s+', 2)[1] } else { '' }

        switch ($verb) {
            { $_ -in 'b','back','q','quit','exit' } { return }
            { $_ -in 'format','fmt' }   { Set-GovFormat $arg }
            'table' { $script:Gov.Format = 'Table' }
            'list'  { $script:Gov.Format = 'List' }
            'json'  { $script:Gov.Format = 'Json' }
            'csv'   { $script:Gov.Format = 'Csv' }
            'grid'  { $script:Gov.Format = 'Grid' }
            'cols'  {
                if ($arg) { $script:Gov.LastAction.Columns = @($arg -split '[,\s]+' | Where-Object { $_ }) }
                else { $script:Gov.LastAction.Columns = $null }
            }
            'top'      { if ($arg -match '^\d+$') { $script:Gov.Top = [int]$arg }; Invoke-GovRefresh -Action $Action }
            'maxitems' { if ($arg -match '^\d+$') { $script:Gov.MaxItems = [int]$arg }; Invoke-GovRefresh -Action $Action }
            'filter'   { $script:Gov.Filter  = if ($arg) { $arg } else { $null }; Invoke-GovRefresh -Action $Action }
            'select'   { $script:Gov.Select  = if ($arg) { $arg } else { $null }; Invoke-GovRefresh -Action $Action }
            'expand'   { $script:Gov.Expand  = if ($arg) { $arg } else { $null }; Invoke-GovRefresh -Action $Action }
            'orderby'  { $script:Gov.OrderBy = if ($arg) { $arg } else { $null }; Invoke-GovRefresh -Action $Action }
            'find'     { Find-GovView $arg }
            'open'     { Show-GovItemDetail $arg }
            { $_ -in 'refresh','r' } { Invoke-GovRefresh -Action $Action }
            'export'   { Export-GovResult -Path $arg }
            'apiver'   { Set-GovApiVersion $arg; Invoke-GovRefresh -Action $Action }
            { $_ -in 'help','?' } { Show-GovHelp; Read-Host '  Press Enter' | Out-Null }
            default {
                Write-Host "  Unknown command '$verb'. Type 'help' or 'b' to go back." -ForegroundColor Yellow
                Read-Host '  Press Enter' | Out-Null
            }
        }
    }
}

function Invoke-GovRefresh {
    param([Parameter(Mandatory)] $Action)
    $uri = Get-GovActionUri -Action $Action
    try {
        $data = Invoke-GovGraph -Uri $uri -MaxItems $script:Gov.MaxItems -Advanced:([bool]($script:Gov.Filter -or $script:Gov.OrderBy))
        $script:Gov.LastResult = @($data)
        $script:Gov.LastView   = @($data)
    }
    catch {
        Write-Host "Refresh failed: $($_.Exception.Message)" -ForegroundColor Red
        Read-Host '  Press Enter' | Out-Null
    }
}

function Find-GovView {
    param([string] $Text)
    if (-not $Text) { $script:Gov.LastView = @($script:Gov.LastResult); return }
    $needle = $Text.ToLower()
    $script:Gov.LastView = @(
        $script:Gov.LastResult | Where-Object {
            $hit = $false
            foreach ($p in $_.PSObject.Properties) {
                if ($p.Name -like '@odata*') { continue }
                $val = "$($p.Value)"
                if ($val -and $val.ToLower().Contains($needle)) { $hit = $true; break }
            }
            $hit
        }
    )
    Write-Host ("  Filtered to {0} of {1} item(s) matching '{2}'." -f $script:Gov.LastView.Count, $script:Gov.LastResult.Count, $Text) -ForegroundColor Green
    Read-Host '  Press Enter' | Out-Null
}

function Show-GovItemDetail {
    param([string] $Arg)
    $view = @($script:Gov.LastView)
    if ($view.Count -eq 0) { return }
    $idx = 0
    if ($Arg -match '^\d+$') { $idx = [int]$Arg } else { $idx = 0 }
    if ($idx -lt 0 -or $idx -ge $view.Count) { Write-Host "  Index out of range (0..$($view.Count-1))." -ForegroundColor Yellow; Read-Host '  Press Enter' | Out-Null; return }
    Clear-GovHost
    Write-Host ("Item [{0}] full detail:" -f $idx) -ForegroundColor White
    Write-GovRule
    ($view[$idx] | ConvertTo-Json -Depth 20) | Out-Host
    Write-GovRule
    Read-Host '  Press Enter to return' | Out-Null
}

#endregion

#region ----------------------------------------------------- Command handlers

function Set-GovFormat {
    param([string] $Name)
    $map = @{ t='Table'; table='Table'; l='List'; list='List'; json='Json'; j='Json'; csv='Csv'; c='Csv'; grid='Grid'; g='Grid' }
    $key = $Name.Trim().ToLower()
    if ($map.ContainsKey($key)) { $script:Gov.Format = $map[$key]; Write-Host "Format set to $($script:Gov.Format)." -ForegroundColor Green }
    else { Write-Host "Unknown format '$Name'. Use Table|List|Json|Csv|Grid." -ForegroundColor Yellow }
}

function Set-GovApiVersion {
    param([string] $Name)
    $key = $Name.Trim().ToLower()
    switch ($key) {
        'beta' { $script:Gov.ApiVersion = 'beta' }
        'v1.0' { $script:Gov.ApiVersion = 'v1.0' }
        'v1'   { $script:Gov.ApiVersion = 'v1.0' }
        'auto' { $script:Gov.ApiVersion = 'auto' }
        default { Write-Host "Use: apiver <auto|v1.0|beta>" -ForegroundColor Yellow; return }
    }
    Write-Host "Graph API version mode: $($script:Gov.ApiVersion)." -ForegroundColor Green
}

function Show-GovHelp {
    Clear-GovHost
    Show-GovBanner
    Write-Host '  HELP — global & result commands' -ForegroundColor White
    Write-GovRule
    $rows = @(
        'Navigation     <number>      Select a menu item / action'
        '               b / back      Go up one menu level'
        '               q / quit      Exit the tool'
        ''
        'Account        L             Open the Login / Account menu'
        '               who           Show current Graph context + scopes'
        '               logout        Sign out of Graph'
        ''
        'Maintenance    P             Prerequisites menu (install missing modules)'
        '               V / about     About: version, API dates, update status'
        '               U / update    Check GitHub and apply an update if available'
        ''
        'Output         format <fmt>  Table | List | Json | Csv | Grid (alias: fmt, or table/list/json/csv/grid)'
        '               cols a,b,c    Set table columns for the current result (blank = auto)'
        '               find <text>   Client-side filter the displayed result'
        '               open <n>      Show full JSON detail of result item n'
        ''
        'Query          top <n>       Server-side page size ($top); 0 = omit'
        '               maxitems <n>  Max items fetched across pages'
        '               filter <odata>  OData $filter (e.g. status eq ''active'')'
        '               select <p,p>  OData $select   expand <rel>  OData $expand   orderby <p>'
        '               apiver <m>    auto | v1.0 | beta   (target latest/preview APIs)'
        '               refresh / r   Re-run the action with current query options'
        ''
        'Export         export        Auto-name file in export dir (JSON/CSV)'
        '               export <path> Export to a specific .json/.csv/.txt file'
        '               export clip   Copy result JSON to the clipboard'
    )
    foreach ($r in $rows) { Write-Host "  $r" -ForegroundColor Gray }
    Write-GovRule
}

# Handles commands available from the MENU prompt. Returns: 'quit' | 'back' | 'handled' | 'none'
function Invoke-GovMenuCommand {
    param([string] $Input)
    $in   = $Input.Trim()
    $verb = ($in -split '\s+', 2)[0].ToLower()
    $arg  = if ($in -match '\s') { ($in -split '\s+', 2)[1] } else { '' }

    switch ($verb) {
        { $_ -in 'q','quit','exit' } { return 'quit' }
        { $_ -in 'b','back' }        { return 'back' }
        { $_ -in 'l','login' }       { Show-GovLoginMenu; return 'handled' }
        { $_ -in 'p','prereq','prereqs' } { Show-GovPrereqMenu; return 'handled' }
        { $_ -in 'v','about','version' }  { Show-GovAboutMenu; return 'handled' }
        { $_ -in 'u','update' }      {
            if (-not $script:GovUpdate -or -not $script:GovUpdate.Configured) { [void](Get-GovUpdateInfo) }
            if ($script:GovUpdate.Configured -and $script:GovUpdate.UpdateAvailable) { Invoke-GovSelfUpdate; if (-not $script:GovRequestExit) { Read-Host '  Press Enter' | Out-Null } }
            else { Show-GovAboutMenu }
            return 'handled'
        }
        'logout'                     { Disconnect-Gov; Read-Host '  Press Enter' | Out-Null; return 'handled' }
        { $_ -in 'who','whoami' }    { Show-GovWhoAmI; Read-Host '  Press Enter' | Out-Null; return 'handled' }
        'scopes'                     { Write-Host ('  ' + ($script:Gov.Scopes -join "`n  ")) -ForegroundColor DarkGray; Read-Host '  Press Enter' | Out-Null; return 'handled' }
        { $_ -in 'format','fmt' }    { Set-GovFormat $arg; return 'handled' }
        'apiver'                     { Set-GovApiVersion $arg; return 'handled' }
        'maxitems'                   { if ($arg -match '^\d+$') { $script:Gov.MaxItems = [int]$arg; Write-Host "MaxItems=$arg" -ForegroundColor Green }; return 'handled' }
        'top'                        { if ($arg -match '^\d+$') { $script:Gov.Top = [int]$arg; Write-Host "Top=$arg" -ForegroundColor Green }; return 'handled' }
        'export'                     { Export-GovResult -Path $arg; Read-Host '  Press Enter' | Out-Null; return 'handled' }
        'exportdir'                  { if ($arg -and (Test-Path $arg)) { $script:Gov.ExportDir = (Resolve-Path $arg).Path; Write-Host "Export dir set." -ForegroundColor Green }; return 'handled' }
        { $_ -in 'help','?','h' }    { Show-GovHelp; Read-Host '  Press Enter' | Out-Null; return 'handled' }
        { $_ -in 'cls','clear' }     { return 'handled' }
        default                      { return 'none' }
    }
}

#endregion

#region ------------------------------------------------------------ Main loop

function Show-GovMenuItems {
    param([Parameter(Mandatory)] $Menu, [string] $Breadcrumb)
    Write-Host ("  {0}" -f $Breadcrumb) -ForegroundColor White
    Write-GovRule
    $i = 1
    foreach ($item in $Menu.Items) {
        $tag = if ($item.Type -eq 'menu') { '>' } else { ' ' }
        Write-Host '   ' -NoNewline
        Write-Host ('{0,2}.' -f $i) -ForegroundColor Yellow -NoNewline
        Write-Host (' {0} {1}' -f $tag, $item.Title)
        $i++
    }
    Write-GovRule
    Write-Host '    ' -NoNewline; Write-Host 'L.' -ForegroundColor Cyan -NoNewline; Write-Host ' Login / Account        ' -NoNewline
    Write-Host 'F.' -ForegroundColor Cyan -NoNewline; Write-Host ' Format / settings      ' -NoNewline
    Write-Host 'C.' -ForegroundColor Cyan -NoNewline; Write-Host ' Custom Graph query'
    Write-Host '    ' -NoNewline; Write-Host 'P.' -ForegroundColor Cyan -NoNewline; Write-Host ' Prerequisites / install' -NoNewline
    Write-Host 'V.' -ForegroundColor Cyan -NoNewline; Write-Host ' About / version       ' -NoNewline
    Write-Host 'U.' -ForegroundColor Cyan -NoNewline; Write-Host ' Update from GitHub'
    Write-Host '    ' -NoNewline; Write-Host '?.' -ForegroundColor Cyan -NoNewline; Write-Host ' Help                   ' -NoNewline
    if ($Menu -ne $script:Gov.Menu) { Write-Host 'B.' -ForegroundColor Cyan -NoNewline; Write-Host ' Back                   ' -NoNewline }
    else { Write-Host '                          ' -NoNewline }
    Write-Host 'Q.' -ForegroundColor Cyan -NoNewline; Write-Host ' Quit'
    Write-GovRule
}

function Show-GovSettingsMenu {
    Clear-GovHost
    Show-GovBanner
    Write-Host '  OUTPUT FORMAT / SETTINGS' -ForegroundColor White
    Write-GovRule
    Write-Host ("   Current format : {0}" -f $script:Gov.Format)
    Write-Host ("   API version    : {0}" -f $script:Gov.ApiVersion)
    Write-Host ("   Top (page)     : {0}" -f $script:Gov.Top)
    Write-Host ("   Max items      : {0}" -f $script:Gov.MaxItems)
    Write-Host ("   Export dir     : {0}" -f $script:Gov.ExportDir)
    Write-GovRule
    Write-Host '   Commands: format <t|l|json|csv|grid> | apiver <auto|v1.0|beta> | top N | maxitems N | exportdir <path> | b'
    Write-GovRule
    $in = Read-Host '   setting>'
    $r = Invoke-GovMenuCommand $in
    return
}

function Invoke-GovCustomQuery {
    if (-not (Get-GovContext)) {
        Write-Host 'Sign in first (L).' -ForegroundColor Red; Read-Host '  Press Enter' | Out-Null; return
    }
    Clear-GovHost
    Show-GovBanner
    Write-Host '  CUSTOM GRAPH QUERY' -ForegroundColor White
    Write-GovRule
    Write-Host '   Enter a relative Graph path, e.g.:' -ForegroundColor DarkGray
    Write-Host '     identityGovernance/entitlementManagement/accessPackages' -ForegroundColor DarkGray
    Write-Host '     roleManagement/directory/roleDefinitions' -ForegroundColor DarkGray
    Write-GovRule
    $path = Read-Host '   path (b to cancel)'
    if (-not $path -or $path.Trim().ToLower() -eq 'b') { return }
    $action = New-GovAction -Title 'Custom query' -Path $path.Trim()
    Invoke-GovAction -Action $action
}

function Start-GovTui {
    [CmdletBinding()] param()

    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
    Initialize-GovState
    Import-GovConfig
    [void](Update-GovPrereqQuick)
    if (-not $script:GovNoUpdateCheck -and (Test-GovRepoConfigured)) {
        try { [void](Get-GovUpdateInfo -Quiet) } catch { }
    }

    $stack = [System.Collections.Generic.Stack[object]]::new()
    $stack.Push($script:Gov.Menu)

    :main while (-not $script:GovRequestExit) {
        $current = $stack.Peek()
        Clear-GovHost
        Show-GovBanner
        Show-GovMenuItems -Menu $current -Breadcrumb (Get-GovBreadcrumb $stack)

        $sel = (Read-Host ("  [{0}|{1}] Select #, command, or L/F/C/P/V/U/?/B/Q" -f $script:Gov.Format, (Get-GovEffectiveVersion))).Trim()
        if ($sel -eq '') { continue main }

        # Single-letter shortcuts (labeled break/continue target the while loop,
        # because an unlabeled break/continue inside a switch only affects the switch).
        switch ($sel.ToLower()) {
            'l' { Show-GovLoginMenu; continue main }
            'f' { Show-GovSettingsMenu; continue main }
            'c' { Invoke-GovCustomQuery; continue main }
            'p' { Show-GovPrereqMenu; continue main }
            'v' { Show-GovAboutMenu; continue main }
            'u' {
                if (-not $script:GovUpdate -or -not $script:GovUpdate.Configured) { [void](Get-GovUpdateInfo) }
                if ($script:GovUpdate.Configured -and $script:GovUpdate.UpdateAvailable) { Invoke-GovSelfUpdate; if (-not $script:GovRequestExit) { Read-Host '  Press Enter' | Out-Null } }
                else { Show-GovAboutMenu }
                continue main
            }
            { $_ -in '?','help','h' } { Show-GovHelp; Read-Host '  Press Enter' | Out-Null; continue main }
            'q' { break main }
            'b' { if ($stack.Count -gt 1) { [void]$stack.Pop() }; continue main }
        }

        if ($sel -match '^\d+$') {
            $idx = [int]$sel - 1
            if ($idx -ge 0 -and $idx -lt $current.Items.Count) {
                $item = $current.Items[$idx]
                if ($item.Type -eq 'menu') { $stack.Push($item) }
                else { Invoke-GovAction -Action $item }
            }
            else { Write-Host '  Invalid selection.' -ForegroundColor Yellow; Start-Sleep -Milliseconds 700 }
            continue main
        }

        # Otherwise treat as a global command
        $r = Invoke-GovMenuCommand $sel
        switch ($r) {
            'quit' { break main }
            'back' { if ($stack.Count -gt 1) { [void]$stack.Pop() } }
            'none' { Write-Host "  Unknown input '$sel'. Type ? for help." -ForegroundColor Yellow; Start-Sleep -Milliseconds 800 }
            default { }
        }
    }

    Clear-GovHost
    if ($script:GovRequestExit) { Write-Host 'Exiting for relaunch...' -ForegroundColor Cyan }
    else { Write-Host 'Goodbye.' -ForegroundColor Cyan }
}

#endregion

# Capture launch overrides for the login menu defaults.
$script:GovTenant        = $TenantId
$script:GovScopeOverride = $Scopes

if (-not $LoadOnly) {
    if ($UseDeviceCode) {
        Initialize-GovState
        Import-GovConfig
        if (Confirm-GovModule -Quiet) { Connect-Gov -Mode DeviceCode -TenantId $TenantId -Scopes $Scopes }
        else { Write-Host "Microsoft.Graph.Authentication missing - launch the tool and press 'P' to install it." -ForegroundColor Yellow }
    }
    Start-GovTui
}
