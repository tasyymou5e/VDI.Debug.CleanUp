#Requires -Version 7.0
<#
.SYNOPSIS
    Menu-driven GPO / Administrative Templates settings auditor. Runs one report or all reports,
    writing one color-coded HTML file per GPO location.
.DESCRIPTION
    Mirrors the design and reporting style of Get-FSLogixSettings.ps1. Each report corresponds to one
    Administrative Templates node under Local Computer Policy:

        Computer Configuration > Administrative Templates
            FSLogix, Horizon Blast, Omnissa DEM, Omnissa Horizon Agent Configuration,
            Omnissa Horizon Client Configuration, OneDrive, Start Menu and Taskbar
        User Configuration > Administrative Templates
            Horizon Blast, Microsoft Edge, Microsoft Teams, Omnissa DEM,
            Omnissa Horizon Agent Configuration, Omnissa Horizon Client Configuration,
            OneDrive, Outlook For Windows, Start Menu and Taskbar

    For each report the script first VALIDATES the underlying registry location(s). If no location
    exists the report still generates and clearly states that the GPO node is not configured. When
    "Run ALL" is chosen, each report is generated as its own separate, location-titled HTML file.

    Because these ADMX templates are delivered by multiple vendors (and VMware was rebranded to
    Omnissa), each report defines several candidate registry roots — legacy VMware and current
    Omnissa paths — and reports whichever exist. All values under the validated roots are read
    recursively so nothing is missed.
.PARAMETER Report
    Run non-interactively. Use a report key (see -List) or 'All'. If omitted, an interactive menu
    is shown.
.PARAMETER List
    Print the available report keys and exit.
.PARAMETER OutputDir
    Folder where HTML reports are written. Defaults to the script folder.
.NOTES
    Run as Administrator on the golden image or VDI session for complete HKLM coverage.
    HKCU (User Configuration) reflects the policy hive of the account running the script.
#>

[CmdletBinding()]
param(
    [string]$Report,
    [switch]$List,
    [string]$OutputDir = $PSScriptRoot
)

Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

if ([string]::IsNullOrWhiteSpace($OutputDir)) { $OutputDir = (Get-Location).Path }

#region --- Report Definitions ---
# Each report:
#   Key       short id used for -Report / menu selection
#   Scope     'Computer' (HKLM) or 'User' (HKCU)
#   Node      friendly Administrative Templates node name
#   GpoPath   full breadcrumb shown as the report title / location
#   RegRoots  ordered list of candidate registry roots to validate + scan recursively
#   Friendly  optional hashtable: ValueName -> @{ Friendly; Description; ValueMap }
#
# Friendly maps enrich well-known values; everything else is reported generically so the
# report is always complete even for undocumented or custom settings.

$onOff = @{0='Disabled';1='Enabled'}

$reportDefs = @(

    #region COMPUTER CONFIGURATION
    [pscustomobject]@{
        Key='C-FSLogix'; Scope='Computer'; Node='FSLogix'
        GpoPath='Local Computer Policy > Computer Configuration > Administrative Templates > FSLogix'
        RegRoots=@(
            'HKLM:\SOFTWARE\Policies\FSLogix'
            'HKLM:\SOFTWARE\FSLogix'
        )
        Friendly=@{
            'Enabled'                 = @{ Friendly='Container Enabled';            Description='Enables the FSLogix Profile / ODFC container.'; ValueMap=$onOff }
            'VHDLocations'            = @{ Friendly='VHD Locations';                Description='UNC path(s) where container VHD(X) files are stored.'; ValueMap=$null }
            'CCDLocations'            = @{ Friendly='Cloud Cache Locations';        Description='Cloud Cache provider strings (precede VHDLocations).'; ValueMap=$null }
            'SizeInMBs'               = @{ Friendly='VHD Size (MB)';                Description='Default container size in MB.'; ValueMap=$null }
            'VolumeType'              = @{ Friendly='VHD Volume Type';              Description='Container disk format.'; ValueMap=@{0='VHD';1='VHDX'} }
            'IsDynamic'               = @{ Friendly='Dynamic VHD';                  Description='Disk grows dynamically vs. pre-allocated.'; ValueMap=@{0='Fixed';1='Dynamic'} }
            'ProfileType'             = @{ Friendly='Profile Type';                 Description='How the container is mounted.'; ValueMap=@{0='Normal (RW)';1='RO template';2='RO + RW local';3='RO + RW diff disk'} }
            'FlipFlopProfileDirectoryName' = @{ Friendly='Flip Flop Directory Name'; Description='SamAccountName_SID folder naming.'; ValueMap=$onOff }
            'DeleteLocalProfileWhenVHDShouldApply' = @{ Friendly='Delete Local Profile When VHD Applies'; Description='Removes local profile if a container should apply.'; ValueMap=$onOff }
            'LoggingLevel'            = @{ Friendly='Logging Level';                Description='FSLogix log verbosity.'; ValueMap=@{0='Errors';1='Errors+Warn';2='Errors+Warn+Info';3='Verbose'} }
            'LogDir'                  = @{ Friendly='Log Directory';                Description='Path for FSLogix logs.'; ValueMap=$null }
        }
    }
    [pscustomobject]@{
        Key='C-Blast'; Scope='Computer'; Node='Horizon Blast'
        GpoPath='Local Computer Policy > Computer Configuration > Administrative Templates > Horizon Blast'
        RegRoots=@(
            'HKLM:\SOFTWARE\Policies\Omnissa\Horizon\Blast'
            'HKLM:\SOFTWARE\Policies\VMware, Inc.\VMware Blast'
        )
        Friendly=@{}
    }
    [pscustomobject]@{
        Key='C-DEM'; Scope='Computer'; Node='Omnissa DEM'
        GpoPath='Local Computer Policy > Computer Configuration > Administrative Templates > Omnissa DEM'
        RegRoots=@(
            'HKLM:\SOFTWARE\Policies\Omnissa\DEM'
            'HKLM:\SOFTWARE\Policies\VMware DEM'
            'HKLM:\SOFTWARE\Policies\VMware UEM'
        )
        Friendly=@{}
    }
    [pscustomobject]@{
        Key='C-Agent'; Scope='Computer'; Node='Omnissa Horizon Agent Configuration'
        GpoPath='Local Computer Policy > Computer Configuration > Administrative Templates > Omnissa Horizon Agent Configuration'
        RegRoots=@(
            'HKLM:\SOFTWARE\Policies\Omnissa\Horizon\Agent'
            'HKLM:\SOFTWARE\Policies\VMware, Inc.\VMware VDM\Agent\Configuration'
            'HKLM:\SOFTWARE\Policies\VMware, Inc.\VMware VDM\Agent'
        )
        Friendly=@{}
    }
    [pscustomobject]@{
        Key='C-Client'; Scope='Computer'; Node='Omnissa Horizon Client Configuration'
        GpoPath='Local Computer Policy > Computer Configuration > Administrative Templates > Omnissa Horizon Client Configuration'
        RegRoots=@(
            'HKLM:\SOFTWARE\Policies\Omnissa\Horizon\Client'
            'HKLM:\SOFTWARE\Policies\VMware, Inc.\VMware VDM\Client'
        )
        Friendly=@{}
    }
    [pscustomobject]@{
        Key='C-OneDrive'; Scope='Computer'; Node='OneDrive'
        GpoPath='Local Computer Policy > Computer Configuration > Administrative Templates > OneDrive'
        RegRoots=@(
            'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'
        )
        Friendly=@{
            'SilentAccountConfig'        = @{ Friendly='Silently Configure Account';        Description='Sign in OneDrive silently with Windows credentials.'; ValueMap=$onOff }
            'KFMSilentOptIn'             = @{ Friendly='KFM Silent Opt-In (Tenant ID)';     Description='Known Folder Move tenant GUID for silent redirection.'; ValueMap=$null }
            'FilesOnDemandEnabled'       = @{ Friendly='Files On-Demand Enabled';           Description='Enables OneDrive Files On-Demand.'; ValueMap=$onOff }
            'DisablePersonalSync'        = @{ Friendly='Block Personal Accounts';            Description='Prevents syncing personal OneDrive accounts.'; ValueMap=$onOff }
        }
    }
    [pscustomobject]@{
        Key='C-StartMenu'; Scope='Computer'; Node='Start Menu and Taskbar'
        GpoPath='Local Computer Policy > Computer Configuration > Administrative Templates > Start Menu and Taskbar'
        RegRoots=@(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer'
            'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer'
        )
        Friendly=@{}
    }
    #endregion

    #region USER CONFIGURATION
    [pscustomobject]@{
        Key='U-Blast'; Scope='User'; Node='Horizon Blast'
        GpoPath='Local Computer Policy > User Configuration > Administrative Templates > Horizon Blast'
        RegRoots=@(
            'HKCU:\SOFTWARE\Policies\Omnissa\Horizon\Blast'
            'HKCU:\SOFTWARE\Policies\VMware, Inc.\VMware Blast'
        )
        Friendly=@{}
    }
    [pscustomobject]@{
        Key='U-Edge'; Scope='User'; Node='Microsoft Edge'
        GpoPath='Local Computer Policy > User Configuration > Administrative Templates > Microsoft Edge'
        RegRoots=@(
            'HKCU:\SOFTWARE\Policies\Microsoft\Edge'
        )
        Friendly=@{
            'HomepageLocation'           = @{ Friendly='Home Page URL';                 Description='Configured browser home page.'; ValueMap=$null }
            'RestoreOnStartup'           = @{ Friendly='Startup Behavior';              Description='What Edge opens on launch.'; ValueMap=@{1='Restore last session';4='Open list of URLs';5='Open New Tab page'} }
            'BackgroundModeEnabled'      = @{ Friendly='Background Mode';                Description='Keeps Edge running in background after close.'; ValueMap=$onOff }
        }
    }
    [pscustomobject]@{
        Key='U-Teams'; Scope='User'; Node='Microsoft Teams'
        GpoPath='Local Computer Policy > User Configuration > Administrative Templates > Microsoft Teams'
        RegRoots=@(
            'HKCU:\SOFTWARE\Policies\Microsoft\Teams'
            'HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Teams'
        )
        Friendly=@{}
    }
    [pscustomobject]@{
        Key='U-DEM'; Scope='User'; Node='Omnissa DEM'
        GpoPath='Local Computer Policy > User Configuration > Administrative Templates > Omnissa DEM'
        RegRoots=@(
            'HKCU:\SOFTWARE\Policies\Omnissa\DEM'
            'HKCU:\SOFTWARE\Policies\VMware DEM'
            'HKCU:\SOFTWARE\Policies\VMware UEM'
        )
        Friendly=@{}
    }
    [pscustomobject]@{
        Key='U-Agent'; Scope='User'; Node='Omnissa Horizon Agent Configuration'
        GpoPath='Local Computer Policy > User Configuration > Administrative Templates > Omnissa Horizon Agent Configuration'
        RegRoots=@(
            'HKCU:\SOFTWARE\Policies\Omnissa\Horizon\Agent'
            'HKCU:\SOFTWARE\Policies\VMware, Inc.\VMware VDM\Agent\Configuration'
            'HKCU:\SOFTWARE\Policies\VMware, Inc.\VMware VDM\Agent'
        )
        Friendly=@{}
    }
    [pscustomobject]@{
        Key='U-Client'; Scope='User'; Node='Omnissa Horizon Client Configuration'
        GpoPath='Local Computer Policy > User Configuration > Administrative Templates > Omnissa Horizon Client Configuration'
        RegRoots=@(
            'HKCU:\SOFTWARE\Policies\Omnissa\Horizon\Client'
            'HKCU:\SOFTWARE\Policies\VMware, Inc.\VMware VDM\Client'
        )
        Friendly=@{}
    }
    [pscustomobject]@{
        Key='U-OneDrive'; Scope='User'; Node='OneDrive'
        GpoPath='Local Computer Policy > User Configuration > Administrative Templates > OneDrive'
        RegRoots=@(
            'HKCU:\SOFTWARE\Policies\Microsoft\OneDrive'
        )
        Friendly=@{}
    }
    [pscustomobject]@{
        Key='U-Outlook'; Scope='User'; Node='Outlook For Windows'
        GpoPath='Local Computer Policy > User Configuration > Administrative Templates > Outlook For Windows'
        RegRoots=@(
            'HKCU:\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook'
            'HKCU:\SOFTWARE\Policies\Microsoft\OutlookForWindows'
            'HKCU:\SOFTWARE\Policies\Microsoft\Office\Outlook'
        )
        Friendly=@{}
    }
    [pscustomobject]@{
        Key='U-StartMenu'; Scope='User'; Node='Start Menu and Taskbar'
        GpoPath='Local Computer Policy > User Configuration > Administrative Templates > Start Menu and Taskbar'
        RegRoots=@(
            'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer'
            'HKCU:\SOFTWARE\Policies\Microsoft\Windows\Explorer'
        )
        Friendly=@{}
    }
    #endregion
)
#endregion

#region --- Helpers ---

function ConvertTo-RegDisplayPath {
    param([string]$Path)
    ($Path -replace '^Microsoft\.PowerShell\.Core\\Registry::', '' `
           -replace '^HKEY_LOCAL_MACHINE', 'HKLM:' `
           -replace '^HKEY_CURRENT_USER',  'HKCU:' `
           -replace '^HKEY_USERS',         'HKU:')
}

function Format-RegValue {
    param($Raw, $Kind, $Map)
    if ($null -eq $Raw) { return 'Not Configured' }

    # Friendly value-map decode (DWORD/QWORD/String)
    if ($Map) {
        $intKey = 0
        if ([int]::TryParse("$Raw", [ref]$intKey) -and $Map.ContainsKey($intKey)) {
            return "$($Map[$intKey])  ($Raw)"
        } elseif ($Map.ContainsKey("$Raw")) {
            return "$($Map["$Raw"])"
        }
    }

    switch ("$Kind") {
        'Binary'      { return (($Raw | ForEach-Object { '{0:X2}' -f $_ }) -join ' ') }
        'MultiString' { return (@($Raw) -join ' ; ') }
        'DWord'       { return "$Raw" }
        'QWord'       { return "$Raw" }
        default       { return "$Raw" }
    }
}

# Recursively read every value under a registry root.
function Read-RegRoot {
    param([string]$Root)

    $items = [System.Collections.Generic.List[pscustomobject]]::new()
    if (-not (Test-Path $Root)) { return $items }

    $keyPaths = @($Root)
    $keyPaths += (Get-ChildItem -Path $Root -Recurse -ErrorAction SilentlyContinue |
                  Select-Object -ExpandProperty PSPath)

    foreach ($kp in $keyPaths) {
        $regKey = $null
        try { $regKey = Get-Item -LiteralPath $kp -ErrorAction Stop } catch { continue }
        $displayKey = ConvertTo-RegDisplayPath $regKey.Name

        foreach ($vName in $regKey.GetValueNames()) {
            $shown = if ([string]::IsNullOrEmpty($vName)) { '(Default)' } else { $vName }
            $kind  = $regKey.GetValueKind($vName)
            $raw   = $regKey.GetValue($vName, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            $items.Add([pscustomobject]@{
                KeyPath   = $displayKey
                ValueName = $shown
                Raw       = $raw
                Kind      = $kind
            })
        }
    }
    return $items
}

function Get-StatusBadge {
    param([string]$Status)
    switch ($Status) {
        'Policy'        { '<span style="background:#1a6e9e;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;white-space:nowrap">POLICY (GPO)</span>' }
        'Direct'        { '<span style="background:#5a7a2e;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;white-space:nowrap">DIRECT (REG)</span>' }
        'Known'         { '<span style="background:#1a6e9e;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;white-space:nowrap">DOCUMENTED</span>' }
        'Extra'         { '<span style="background:#7b4f9e;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;white-space:nowrap">VALUE SET</span>' }
        'NotConfigured' { '<span style="background:#aaa;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;white-space:nowrap">NOT SET</span>' }
        default         { $Status }
    }
}

function Get-RowBg {
    param([string]$Status)
    switch ($Status) {
        'Policy'        { '#e8f4fd' }
        'Direct'        { '#edf7e6' }
        'Known'         { '#e8f4fd' }
        'Extra'         { '#f5eef8' }
        'NotConfigured' { '#f8f9fa' }
        default         { '#fff' }
    }
}

#endregion

#region --- Report Generation ---

function New-GpoReport {
    param([Parameter(Mandatory)] $Def, [Parameter(Mandatory)][string]$OutDir)

    $enc = { param($s) [System.Web.HttpUtility]::HtmlEncode("$s") }

    Write-Host "`n=== $($Def.Node) [$($Def.Scope) Configuration] ===" -ForegroundColor Cyan
    Write-Host "    Location: $($Def.GpoPath)" -ForegroundColor Gray

    # 1) VALIDATE LOCATION FIRST
    Write-Host "[1/3] Validating GPO registry location(s)..." -ForegroundColor Cyan
    $rootStatus = foreach ($root in $Def.RegRoots) {
        $exists = Test-Path $root
        [pscustomobject]@{ Root=$root; Exists=$exists }
        $tag = if ($exists) { 'FOUND   ' } else { 'missing ' }
        $col = if ($exists) { 'Green' } else { 'DarkGray' }
        Write-Host "      [$tag] $root" -ForegroundColor $col
    }
    $validRoots = @($rootStatus | Where-Object Exists | Select-Object -ExpandProperty Root)

    # 2) READ VALUES
    Write-Host "[2/3] Reading policy values..." -ForegroundColor Cyan
    $rows = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($root in $validRoots) {
        $isPolicy = $root -match '\\Policies\\'
        foreach ($v in (Read-RegRoot -Root $root)) {
            $info = $null
            if ($Def.Friendly -and $Def.Friendly.ContainsKey($v.ValueName)) {
                $info = $Def.Friendly[$v.ValueName]
            }
            $friendlyName = if ($info) { $info.Friendly } else { $v.ValueName }
            $desc         = if ($info) { $info.Description } else { 'Policy value present in the registry under this Administrative Templates node.' }
            $map          = if ($info) { $info.ValueMap } else { $null }
            $status       = if ($info) { 'Known' } elseif ($isPolicy) { 'Policy' } else { 'Direct' }

            $rows.Add(@{
                Status      = $status
                Friendly    = $friendlyName
                Value       = (Format-RegValue -Raw $v.Raw -Kind $v.Kind -Map $map)
                Kind        = "$($v.Kind)"
                Reg         = "$($v.KeyPath)\$($v.ValueName)"
                Description = $desc
            })
        }
    }

    Write-Host "      $($rows.Count) value(s) read across $($validRoots.Count) location(s)." -ForegroundColor Gray

    # 2b) DETECT POLICY CONFLICTS
    # A conflict exists when the same ValueName appears under BOTH a Policies path
    # and a non-Policies (direct) path with different values. The Policies key wins
    # at runtime, but having both present with different values indicates a stale
    # direct write that may confuse administrators.
    $conflicts = [System.Collections.Generic.List[hashtable]]::new()

    $policyRows = $rows | Where-Object { $_.Status -in 'Policy','Known' }
    $directRows = $rows | Where-Object { $_.Status -eq 'Direct' }

    foreach ($pRow in $policyRows) {
        # Extract the ValueName from the Reg path (everything after the last \)
        $vName = $pRow.Reg -replace '^.*\\', ''
        $matchDirect = $directRows | Where-Object { ($_.Reg -replace '^.*\\', '') -eq $vName }
        foreach ($dRow in $matchDirect) {
            if ($pRow.Value -ne $dRow.Value) {
                $conflicts.Add(@{
                    ValueName   = $vName
                    PolicyPath  = $pRow.Reg
                    PolicyValue = $pRow.Value
                    DirectPath  = $dRow.Reg
                    DirectValue = $dRow.Value
                })
                Write-Host "      CONFLICT: $vName — Policy='$($pRow.Value)' vs Direct='$($dRow.Value)'" -ForegroundColor Yellow
            }
        }
    }

    if ($conflicts.Count -gt 0) {
        Write-Host "      $($conflicts.Count) policy conflict(s) detected." -ForegroundColor Yellow
    }

    # 3) BUILD HTML
    Write-Host "[3/3] Generating HTML report..." -ForegroundColor Cyan

    $policyCount   = ($rows | Where-Object { $_.Status -in 'Policy','Known' }).Count
    $directCount   = ($rows | Where-Object { $_.Status -eq 'Direct' }).Count
    $conflictCount = $conflicts.Count
    $totalCount    = $rows.Count

    # Validation table
    $valRows = $rootStatus | ForEach-Object {
        $bg = if ($_.Exists) { '#e8f4fd' } else { '#fff3cd' }
        $badge = if ($_.Exists) {
            '<span style="background:#28a745;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em">FOUND</span>'
        } else {
            '<span style="background:#dc3545;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em">NOT PRESENT</span>'
        }
        "<tr style=`"background:$bg`"><td>$badge</td><td class=`"mono small`">$(& $enc $_.Root)</td></tr>"
    }

    # Settings table (or empty-state note)
    if ($totalCount -gt 0) {
        $settingRows = $rows | ForEach-Object {
            $bg  = Get-RowBg $_.Status
            $bdg = Get-StatusBadge $_.Status
            @"
        <tr style="background:$bg">
          <td>$bdg</td>
          <td><strong>$(& $enc $_.Friendly)</strong><br><span class="desc">$(& $enc $_.Description)</span></td>
          <td class="mono">$(& $enc $_.Value)</td>
          <td class="mono small">$(& $enc $_.Kind)</td>
          <td class="mono small">$(& $enc $_.Reg)</td>
        </tr>
"@
        }
        $settingsHtml = @"
    <table>
      <thead><tr>
        <th style="width:120px">Source</th>
        <th>Setting</th>
        <th style="width:200px">Current Value</th>
        <th style="width:90px">Type</th>
        <th style="width:300px">Registry Path \ Value Name</th>
      </tr></thead>
      <tbody>$($settingRows -join '')</tbody>
    </table>
"@
    }
    else {
        $settingsHtml = @"
    <div class="empty">
      <strong>This GPO node is not configured.</strong><br>
      No policy values were found under the candidate registry location(s) for
      <em>$(& $enc $Def.Node)</em> ($($Def.Scope) Configuration). The settings are at their
      Windows / product default state.
    </div>
"@
    }

    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>GPO Audit — $(& $enc $Def.Node) ($($Def.Scope)) — $(Get-Date -Format 'yyyy-MM-dd HH:mm')</title>
  <style>
    body   { font-family: Segoe UI, Arial, sans-serif; margin: 30px; background: #f4f6f8; color: #222; }
    h1     { color: #1a3a5c; border-bottom: 3px solid #1a3a5c; padding-bottom: 8px; margin-bottom: 4px; }
    .crumb { color:#1a6e9e; font-size:0.95em; margin:2px 0 0 0; }
    h2     { color: #1a6e9e; margin-top: 32px; margin-bottom: 8px; font-size: 1.1em;
             border-left: 4px solid #1a6e9e; padding-left: 10px; }
    .meta  { background:#fff; border-radius:6px; padding:14px 20px; margin:16px 0;
             box-shadow:0 1px 3px rgba(0,0,0,.1); font-size:0.95em; }
    .summary { display:flex; gap:14px; flex-wrap:wrap; margin:20px 0; }
    .badge { padding:12px 20px; border-radius:8px; color:#fff; font-weight:bold;
             font-size:1em; min-width:90px; text-align:center; box-shadow:0 1px 3px rgba(0,0,0,.15); }
    table  { width:100%; border-collapse:collapse; background:#fff; border-radius:8px;
             overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,.08); margin-bottom:10px; }
    th     { background:#1a3a5c; color:#fff; padding:9px 12px; text-align:left; font-size:0.88em; }
    td     { padding:8px 12px; border-bottom:1px solid #e3e8ee; vertical-align:top; font-size:0.88em; }
    tr:last-child td { border-bottom:none; }
    .mono  { font-family: Consolas, 'Courier New', monospace; font-size:0.85em; word-break:break-all; }
    .small { font-size:0.78em; color:#555; }
    .desc  { font-size:0.82em; color:#666; display:block; margin-top:2px; }
    .note  { background:#e8f4fd; border-left:4px solid #1a6e9e; padding:10px 16px;
             border-radius:0 6px 6px 0; margin:16px 0; font-size:0.9em; }
    .empty { background:#fff3cd; border-left:4px solid #e0a800; padding:14px 18px;
             border-radius:0 6px 6px 0; margin:16px 0; font-size:0.95em; }
    .legend { display:flex; gap:12px; flex-wrap:wrap; margin:8px 0 20px 0; font-size:0.85em; align-items:center; }
    code   { background:#eef2f6; padding:1px 5px; border-radius:3px; font-size:0.85em; }
  </style>
</head>
<body>

<h1>GPO Settings Audit — $(& $enc $Def.Node)</h1>
<p class="crumb">$(& $enc $Def.GpoPath)</p>

<div class="meta">
  <strong>Machine:</strong> $($env:COMPUTERNAME) &nbsp;|&nbsp;
  <strong>Domain:</strong> $($env:USERDOMAIN) &nbsp;|&nbsp;
  <strong>Run by:</strong> $($env:USERNAME) &nbsp;|&nbsp;
  <strong>Scope:</strong> $($Def.Scope) Configuration &nbsp;|&nbsp;
  <strong>Date:</strong> $stamp &nbsp;|&nbsp;
  <strong>OS:</strong> $((Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption)
</div>

<h2 style="border:none;margin-top:10px">Summary</h2>
<div class="summary">
  <div class="badge" style="background:#1a6e9e">Policy / Documented<br>$policyCount</div>
  <div class="badge" style="background:#5a7a2e">Direct (Reg)<br>$directCount</div>
  <div class="badge" style="background:#1a3a5c">Total Values<br>$totalCount</div>
  $(if ($conflictCount -gt 0) { "<div class=`"badge`" style=`"background:#c0392b`">&#9888; Conflicts<br>$conflictCount</div>" })
</div>

<div class="legend">
  <strong>Legend:</strong>
  <span style="background:#1a6e9e;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.8em">POLICY (GPO)</span> Under a <code>\Policies\</code> key — enforced &nbsp;
  <span style="background:#1a6e9e;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.8em">DOCUMENTED</span> Known ADMX setting with friendly mapping &nbsp;
  <span style="background:#5a7a2e;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.8em">DIRECT (REG)</span> Set outside a Policies key &nbsp;
  <span style="background:#aaa;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.8em">NOT SET</span> Node not configured
</div>

<h2>Location Validation</h2>
<p style="color:#666;font-size:0.9em">Candidate registry roots backing this Administrative Templates node (legacy VMware and current Omnissa paths are both checked where applicable):</p>
<table>
  <thead><tr><th style="width:130px">Status</th><th>Registry Root</th></tr></thead>
  <tbody>$($valRows -join '')</tbody>
</table>

$(if ($conflictCount -gt 0) {
    $conflictRows = $conflicts | ForEach-Object {
        @"
        <tr style="background:#fdf2f2">
          <td class="mono small">$(& $enc $_.ValueName)</td>
          <td class="mono small">$(& $enc $_.PolicyPath)</td>
          <td class="mono small" style="color:#c0392b"><strong>$(& $enc $_.PolicyValue)</strong> (wins)</td>
          <td class="mono small">$(& $enc $_.DirectPath)</td>
          <td class="mono small" style="color:#888">$(& $enc $_.DirectValue) (ignored)</td>
        </tr>
"@
    }
    @"
<h2 style="border-left-color:#c0392b;color:#c0392b">&#9888; Policy Conflicts Detected ($conflictCount)</h2>
<div class="note" style="border-left-color:#c0392b;background:#fdf2f2">
  The following settings exist under <strong>both</strong> a <code>\Policies\</code> path (enforced by GPO) and a direct registry path with <strong>different values</strong>.
  The Policies path wins at runtime. The direct-path value is stale and may be confusing. Consider removing the direct-path value to avoid ambiguity.
</div>
<table>
  <thead><tr>
    <th>Value Name</th>
    <th>Policy Path</th>
    <th>Policy Value (active)</th>
    <th>Direct Path</th>
    <th>Direct Value (overridden)</th>
  </tr></thead>
  <tbody>$($conflictRows -join '')</tbody>
</table>
"@
})

<h2>Configured Settings</h2>
$settingsHtml

<p style="margin-top:40px;color:#999;font-size:0.8em">
  Generated by Get-GPOSettings.ps1 &mdash; PowerShell $($PSVersionTable.PSVersion) &mdash; $stamp
</p>

</body>
</html>
"@

    $safeNode = ($Def.Node -replace '[^\w]+', '_').Trim('_')
    $file = Join-Path $OutDir ("GPO_{0}_{1}_{2}.html" -f $Def.Scope, $safeNode, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $html | Out-File -FilePath $file -Encoding UTF8
    Write-Host "      Report saved: $file" -ForegroundColor Green

    return [pscustomobject]@{ Node=$Def.Node; Scope=$Def.Scope; File=$file; Values=$totalCount }
}

#endregion

#region --- Menu / Dispatch ---

if ($List) {
    Write-Host "`nAvailable GPO reports:`n" -ForegroundColor Cyan
    $reportDefs | ForEach-Object {
        '{0,-12} {1,-9} {2}' -f $_.Key, $_.Scope, $_.Node
    }
    Write-Host "`nUse:  -Report <Key>   |   -Report All`n"
    return
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

function Invoke-Selection {
    param([object[]]$Defs, [string]$OutDir)
    $generated = foreach ($d in $Defs) { New-GpoReport -Def $d -OutDir $OutDir }

    Write-Host "`n================ Done ================" -ForegroundColor Cyan
    $generated | ForEach-Object {
        Write-Host ("  [{0,-8}] {1,-38} {2} value(s)" -f $_.Scope, $_.Node, $_.Values) -ForegroundColor Gray
    }
    Write-Host ("  {0} report(s) written to: {1}" -f @($generated).Count, $OutDir) -ForegroundColor Green

    if (@($generated).Count -eq 1) {
        try { Start-Process $generated[0].File } catch { }
    } else {
        try { Start-Process $OutDir } catch { }
    }
}

# Non-interactive path
if ($Report) {
    if ($Report -ieq 'All') {
        Invoke-Selection -Defs $reportDefs -OutDir $OutputDir
    } else {
        $sel = $reportDefs | Where-Object { $_.Key -ieq $Report -or $_.Node -ieq $Report }
        if (-not $sel) {
            Write-Host "Unknown report '$Report'. Use -List to see valid keys." -ForegroundColor Red
            return
        }
        Invoke-Selection -Defs @($sel) -OutDir $OutputDir
    }
    return
}

# Interactive menu
while ($true) {
    Write-Host "`n=====================================================================" -ForegroundColor Cyan
    Write-Host "  GPO / Administrative Templates Report Generator" -ForegroundColor White
    Write-Host "  Local Computer Policy  ->  Administrative Templates" -ForegroundColor DarkGray
    Write-Host "=====================================================================" -ForegroundColor Cyan

    $menu = @()
    $i = 0

    Write-Host "`n  -- Computer Configuration --" -ForegroundColor Yellow
    foreach ($d in ($reportDefs | Where-Object Scope -eq 'Computer')) {
        $i++; $menu += $d
        Write-Host ("   {0,2}. {1}" -f $i, $d.Node)
    }
    Write-Host "`n  -- User Configuration --" -ForegroundColor Yellow
    foreach ($d in ($reportDefs | Where-Object Scope -eq 'User')) {
        $i++; $menu += $d
        Write-Host ("   {0,2}. {1}" -f $i, $d.Node)
    }

    Write-Host ""
    Write-Host "    A. Run ALL reports (one separate HTML per location)" -ForegroundColor Green
    Write-Host "    Q. Quit"
    Write-Host ""
    $choice = Read-Host "Select a number, 'A' for all, or 'Q' to quit"

    if ($choice -ieq 'Q') { break }

    if ($choice -ieq 'A') {
        Invoke-Selection -Defs $reportDefs -OutDir $OutputDir
        continue
    }

    $n = 0
    if ([int]::TryParse($choice, [ref]$n) -and $n -ge 1 -and $n -le $menu.Count) {
        Invoke-Selection -Defs @($menu[$n-1]) -OutDir $OutputDir
    } else {
        Write-Host "Invalid selection." -ForegroundColor Red
    }
}

#endregion
