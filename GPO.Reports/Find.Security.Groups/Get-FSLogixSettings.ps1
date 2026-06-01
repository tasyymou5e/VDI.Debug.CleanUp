#Requires -Version 7.0
<#
.SYNOPSIS
    Documents all FSLogix Local Computer Policy settings and writes a detailed HTML report.
.DESCRIPTION
    Reads FSLogix configuration from both policy-managed registry keys (HKLM:\SOFTWARE\Policies\FSLogix)
    and direct configuration keys (HKLM:\SOFTWARE\FSLogix). Maps every value to its human-readable
    setting name, expected values, and current status. Outputs a color-coded HTML report.
.NOTES
    Run as Administrator on the golden image or VDI session.
#>

[CmdletBinding()]
param(
    [string]$ReportPath = "$PSScriptRoot\FSLogix_Settings_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
)

Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

#region --- Registry Value Definitions ---
# Each entry: RegPath, ValueName, FriendlyName, Section, Description, ValueMap (hashtable of value->meaning)

$fslogixDefs = @(

    #region PROFILE CONTAINER
    [pscustomobject]@{ Path='Profiles'; Name='Enabled';                    Section='Profile Container'; Friendly='Profile Container Enabled';                Description='Enables or disables FSLogix Profile Container.';                                       ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='Profiles'; Name='VHDLocations';               Section='Profile Container'; Friendly='VHD Locations';                           Description='Semicolon-separated list of UNC paths where profile VHDs are stored.';               ValueMap=$null }
    [pscustomobject]@{ Path='Profiles'; Name='CCDLocations';               Section='Profile Container'; Friendly='Cloud Cache Locations';                   Description='Cloud Cache VHD provider strings (takes precedence over VHDLocations).';             ValueMap=$null }
    [pscustomobject]@{ Path='Profiles'; Name='SizeInMBs';                  Section='Profile Container'; Friendly='VHD Size (MB)';                          Description='Default size in MB of the profile VHD when first created.';                           ValueMap=$null }
    [pscustomobject]@{ Path='Profiles'; Name='VolumeType';                 Section='Profile Container'; Friendly='VHD Volume Type';                        Description='VHD format to use for profile container.';                                            ValueMap=@{0='VHD';1='VHDX'} }
    [pscustomobject]@{ Path='Profiles'; Name='IsDynamic';                  Section='Profile Container'; Friendly='Dynamic VHD';                            Description='When enabled, VHD grows dynamically rather than pre-allocating full size.';           ValueMap=@{0='Fixed';1='Dynamic'} }
    [pscustomobject]@{ Path='Profiles'; Name='FlipFlopProfileDirectoryName'; Section='Profile Container'; Friendly='Flip Flop Profile Directory Name';    Description='Reverses profile folder naming convention to SamAccountName_SID instead of SID_SamAccountName.'; ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='Profiles'; Name='ProfileType';                Section='Profile Container'; Friendly='Profile Type';                           Description='Defines how FSLogix handles the profile container type.';                             ValueMap=@{0='Normal (RW)';1='Read-only template';2='Read-only (merge RW local)';3='Read-only (merge RW diff disk)'} }
    [pscustomobject]@{ Path='Profiles'; Name='ConcurrentUserSessions';     Section='Profile Container'; Friendly='Concurrent User Sessions';               Description='Allows multiple simultaneous sessions for the same user.';                            ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='Profiles'; Name='DeleteLocalProfileWhenVHDShouldApply'; Section='Profile Container'; Friendly='Delete Local Profile When VHD Applies'; Description='Deletes existing local profile if a VHD profile exists and should be applied.'; ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='Profiles'; Name='PreventLoginWithFailure';    Section='Profile Container'; Friendly='Prevent Login With Failure';             Description='Prevents user login if FSLogix Profile Container fails to attach.';                  ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='Profiles'; Name='PreventLoginWithTempProfile'; Section='Profile Container'; Friendly='Prevent Login With Temp Profile';      Description='Prevents user login if only a temporary profile can be provided.';                   ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='Profiles'; Name='ReAttachIntervalSeconds';    Section='Profile Container'; Friendly='Re-Attach Interval (Seconds)';           Description='How often FSLogix attempts to re-attach a disconnected VHD.';                        ValueMap=$null }
    [pscustomobject]@{ Path='Profiles'; Name='ReAttachRetryCount';         Section='Profile Container'; Friendly='Re-Attach Retry Count';                  Description='Number of re-attach attempts before FSLogix gives up.';                              ValueMap=$null }
    [pscustomobject]@{ Path='Profiles'; Name='RoamSearch';                 Section='Profile Container'; Friendly='Roam Search Database';                   Description='Controls roaming of Windows Search index database.';                                  ValueMap=@{0='Disabled';1='Enabled — multi-user';2='Enabled — single user'} }
    [pscustomobject]@{ Path='Profiles'; Name='RoamRecycleBin';             Section='Profile Container'; Friendly='Roam Recycle Bin';                       Description='Includes the Recycle Bin in the profile container.';                                  ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='Profiles'; Name='SetTempToLocalPath';         Section='Profile Container'; Friendly='Set Temp to Local Path';                 Description='Redirects TEMP/TMP environment variables to a local path instead of the profile.';   ValueMap=@{0='Disabled';1='Temp and TMP';2='Temp only';3='TMP only'} }
    [pscustomobject]@{ Path='Profiles'; Name='LockedRetryCount';           Section='Profile Container'; Friendly='Locked VHD Retry Count';                 Description='Number of times to retry attaching a VHD that is locked by another session.';        ValueMap=$null }
    [pscustomobject]@{ Path='Profiles'; Name='LockedRetryInterval';        Section='Profile Container'; Friendly='Locked VHD Retry Interval (Seconds)';    Description='Seconds to wait between locked VHD retry attempts.';                                  ValueMap=$null }
    [pscustomobject]@{ Path='Profiles'; Name='OutlookCachedMode';          Section='Profile Container'; Friendly='Outlook Cached Mode';                    Description='Forces Outlook to use cached mode — recommended with profile containers.';           ValueMap=@{0='Not configured';1='Enabled'} }
    [pscustomobject]@{ Path='Profiles'; Name='AccessNetworkAsComputerObject'; Section='Profile Container'; Friendly='Access Network as Computer Object';  Description='Allows FSLogix to access the file server using the machine account.';                 ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='Profiles'; Name='CleanupInvalidSessions';     Section='Profile Container'; Friendly='Cleanup Invalid Sessions';               Description='Removes stale VHD locks left by crashed sessions.';                                   ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='Profiles'; Name='KeepLocalDir';               Section='Profile Container'; Friendly='Keep Local Profile Directory';           Description='Retains the local profile folder after VHD detach (for troubleshooting).';           ValueMap=@{0='Delete';1='Keep'} }
    [pscustomobject]@{ Path='Profiles'; Name='MirrorLocalProfilesToCloud';  Section='Profile Container'; Friendly='Mirror Local Profiles to Cloud';        Description='Mirrors local profiles into the cloud container on login.';                           ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='Profiles'; Name='RedirectType';               Section='Profile Container'; Friendly='Redirection Type';                       Description='Controls how FSLogix handles folder redirection within the container.';               ValueMap=@{0='Symlinks';1='Junctions'} }
    [pscustomobject]@{ Path='Profiles'; Name='InstallAppxPackages';        Section='Profile Container'; Friendly='Roam Appx Packages';                     Description='Roams AppX/MSIX app packages in the profile container.';                             ValueMap=@{0='Disabled';1='Enabled'} }
    #endregion

    #region OFFICE 365 CONTAINER (ODFC)
    [pscustomobject]@{ Path='ODFC'; Name='Enabled';                Section='Office 365 Container'; Friendly='ODFC Container Enabled';              Description='Enables or disables the Office 365 Data (ODFC) Container.';                          ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='ODFC'; Name='VHDLocations';           Section='Office 365 Container'; Friendly='VHD Locations';                       Description='Semicolon-separated UNC paths for Office container VHDs.';                           ValueMap=$null }
    [pscustomobject]@{ Path='ODFC'; Name='CCDLocations';           Section='Office 365 Container'; Friendly='Cloud Cache Locations';               Description='Cloud Cache provider strings for Office container.';                                 ValueMap=$null }
    [pscustomobject]@{ Path='ODFC'; Name='SizeInMBs';              Section='Office 365 Container'; Friendly='VHD Size (MB)';                      Description='Default VHD size in MB for the Office container.';                                   ValueMap=$null }
    [pscustomobject]@{ Path='ODFC'; Name='VolumeType';             Section='Office 365 Container'; Friendly='VHD Volume Type';                    Description='VHD format for Office container.';                                                   ValueMap=@{0='VHD';1='VHDX'} }
    [pscustomobject]@{ Path='ODFC'; Name='IsDynamic';              Section='Office 365 Container'; Friendly='Dynamic VHD';                        Description='Dynamic disk growth for Office container.';                                           ValueMap=@{0='Fixed';1='Dynamic'} }
    [pscustomobject]@{ Path='ODFC'; Name='FlipFlopProfileDirectoryName'; Section='Office 365 Container'; Friendly='Flip Flop Directory Name';   Description='Reverses Office container folder naming to SamAccountName_SID.';                    ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='ODFC'; Name='IncludeTeams';           Section='Office 365 Container'; Friendly='Include Teams';                      Description='Includes Microsoft Teams data in the ODFC container.';                               ValueMap=@{0='Excluded';1='Included'} }
    [pscustomobject]@{ Path='ODFC'; Name='IncludeOneDrive';        Section='Office 365 Container'; Friendly='Include OneDrive';                   Description='Includes OneDrive cache in the ODFC container.';                                     ValueMap=@{0='Excluded';1='Included'} }
    [pscustomobject]@{ Path='ODFC'; Name='IncludeSharepoint';      Section='Office 365 Container'; Friendly='Include SharePoint';                 Description='Includes SharePoint cache in the ODFC container.';                                   ValueMap=@{0='Excluded';1='Included'} }
    [pscustomobject]@{ Path='ODFC'; Name='IncludeOutlook';         Section='Office 365 Container'; Friendly='Include Outlook';                    Description='Includes Outlook OST and data in the ODFC container.';                               ValueMap=@{0='Excluded';1='Included'} }
    [pscustomobject]@{ Path='ODFC'; Name='IncludeOutlookPersonalization'; Section='Office 365 Container'; Friendly='Include Outlook Personalization'; Description='Includes Outlook personalization data (signatures, templates) in ODFC.';      ValueMap=@{0='Excluded';1='Included'} }
    [pscustomobject]@{ Path='ODFC'; Name='IncludeOfficeActivation'; Section='Office 365 Container'; Friendly='Include Office Activation';         Description='Includes Office activation tokens in the ODFC container.';                          ValueMap=@{0='Excluded';1='Included'} }
    [pscustomobject]@{ Path='ODFC'; Name='IncludeOneNote';         Section='Office 365 Container'; Friendly='Include OneNote';                    Description='Includes OneNote data in the ODFC container.';                                       ValueMap=@{0='Excluded';1='Included'} }
    [pscustomobject]@{ Path='ODFC'; Name='IncludeOneNote_UWP';     Section='Office 365 Container'; Friendly='Include OneNote UWP';                Description='Includes UWP (Store) version of OneNote in ODFC container.';                         ValueMap=@{0='Excluded';1='Included'} }
    [pscustomobject]@{ Path='ODFC'; Name='PreventLoginWithFailure'; Section='Office 365 Container'; Friendly='Prevent Login With Failure';        Description='Prevents login if ODFC container fails to attach.';                                  ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='ODFC'; Name='AccessNetworkAsComputerObject'; Section='Office 365 Container'; Friendly='Access Network as Computer Object'; Description='Uses machine account for ODFC file server access.';                       ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='ODFC'; Name='LockedRetryCount';        Section='Office 365 Container'; Friendly='Locked VHD Retry Count';            Description='Retries for locked ODFC VHD.';                                                       ValueMap=$null }
    [pscustomobject]@{ Path='ODFC'; Name='LockedRetryInterval';     Section='Office 365 Container'; Friendly='Locked VHD Retry Interval (Sec)';  Description='Seconds between locked ODFC VHD retries.';                                          ValueMap=$null }
    #endregion

    #region LOCAL GROUPS (Include/Exclude)
    [pscustomobject]@{ Path='Profiles'; Name='IncludeListRegistry';    Section='Include / Exclude Groups'; Friendly='Profile Include List (Registry)';    Description='Users/groups in this list will use Profile Container. If empty, all users are included unless in Exclude List.'; ValueMap=$null }
    [pscustomobject]@{ Path='Profiles'; Name='ExcludeListRegistry';    Section='Include / Exclude Groups'; Friendly='Profile Exclude List (Registry)';    Description='Users/groups in this list are excluded from Profile Container.';                ValueMap=$null }
    [pscustomobject]@{ Path='ODFC';     Name='IncludeListRegistry';    Section='Include / Exclude Groups'; Friendly='ODFC Include List (Registry)';       Description='Users/groups in this list will use ODFC Container.';                           ValueMap=$null }
    [pscustomobject]@{ Path='ODFC';     Name='ExcludeListRegistry';    Section='Include / Exclude Groups'; Friendly='ODFC Exclude List (Registry)';       Description='Users/groups excluded from ODFC Container.';                                  ValueMap=$null }
    #endregion

    #region LOGGING
    [pscustomobject]@{ Path='Logging'; Name='LogDir';              Section='Logging'; Friendly='Log Directory';                    Description='Path where FSLogix writes its log files.';                                                          ValueMap=$null }
    [pscustomobject]@{ Path='Logging'; Name='LoggingLevel';        Section='Logging'; Friendly='Logging Level';                   Description='Controls verbosity of FSLogix log output.';                                                          ValueMap=@{0='Only errors';1='Errors and warnings';2='Errors, warnings, info';3='All (verbose)'} }
    [pscustomobject]@{ Path='Logging'; Name='LogFileKeepingPeriod'; Section='Logging'; Friendly='Log File Retention (Days)';     Description='Number of days to retain FSLogix log files before deletion.';                                        ValueMap=$null }
    #endregion

    #region APPS / GENERAL
    [pscustomobject]@{ Path='Apps'; Name='CleanupInvalidSessions';  Section='Apps / General'; Friendly='Cleanup Invalid Sessions';    Description='Globally enables cleanup of stale VHD lock files.';                                             ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='Apps'; Name='VHDCompactDisk';          Section='Apps / General'; Friendly='Compact VHD on Detach';       Description='Compacts the VHD file when the container is detached to reclaim unused space.';                 ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='Apps'; Name='RoamIdentity';            Section='Apps / General'; Friendly='Roam Identity (AAD/ADAL)';    Description='Roams Azure AD / ADAL identity tokens in the profile container.';                              ValueMap=@{0='Disabled';1='Enabled'} }
    [pscustomobject]@{ Path='Apps'; Name='LibraryRegistryPath';     Section='Apps / General'; Friendly='Library Registry Path';       Description='Custom path for FSLogix library registry entries.';                                             ValueMap=$null }
    #endregion
)

#endregion

#region --- Registry Read ---

Write-Host "`n[1/3] Reading FSLogix registry settings..." -ForegroundColor Cyan

# Base registry paths — Policy (GPO-applied) takes precedence over direct
$policyBase = 'HKLM:\SOFTWARE\Policies\FSLogix'
$directBase = 'HKLM:\SOFTWARE\FSLogix'

$results = [System.Collections.Generic.List[hashtable]]::new()

foreach ($def in $fslogixDefs) {

    $policyPath = Join-Path $policyBase $def.Path
    $directPath = Join-Path $directBase $def.Path

    $source = $null
    $rawVal = $null

    # Policy path takes precedence
    if (Test-Path $policyPath) {
        $val = Get-ItemProperty $policyPath -Name $def.Name -ErrorAction SilentlyContinue
        if ($null -ne $val -and $null -ne $val.($def.Name)) {
            $rawVal = $val.($def.Name)
            $source = 'Policy (ADMX/GPO)'
        }
    }

    # Fall back to direct path
    if ($null -eq $rawVal -and (Test-Path $directPath)) {
        $val = Get-ItemProperty $directPath -Name $def.Name -ErrorAction SilentlyContinue
        if ($null -ne $val -and $null -ne $val.($def.Name)) {
            $rawVal = $val.($def.Name)
            $source = 'Direct (Registry)'
        }
    }

    # Resolve value to friendly text
    $friendlyVal = if ($null -eq $rawVal) {
        'Not Configured'
    } elseif ($def.ValueMap) {
        # Try integer key first, then string key
        $intKey = $null
        $intOk  = [int]::TryParse("$rawVal", [ref]$intKey)
        if ($intOk -and $def.ValueMap.ContainsKey($intKey)) {
            "$($def.ValueMap[$intKey])  ($rawVal)"
        } elseif ($def.ValueMap.ContainsKey("$rawVal")) {
            "$($def.ValueMap["$rawVal"])"
        } else {
            "$rawVal"
        }
    } else {
        "$rawVal"
    }

    $status = if ($null -eq $rawVal) { 'NotConfigured' }
              elseif ($source -eq 'Policy (ADMX/GPO)') { 'Policy' }
              else { 'Direct' }

    $results.Add(@{
        Section     = $def.Section
        Friendly    = $def.Friendly
        RegPath     = if ($source -eq 'Policy (ADMX/GPO)') { "$policyBase\$($def.Path)" } else { "$directBase\$($def.Path)" }
        RegName     = $def.Name
        Value       = $friendlyVal
        RawValue    = $rawVal
        Source      = if ($source) { $source } else { '—' }
        Status      = $status
        Description = $def.Description
    })
}

# Also scan for any UNDOCUMENTED values present in the registry (catch custom/unlisted settings)
$extraResults = [System.Collections.Generic.List[hashtable]]::new()
$allSubKeys = @(
    "$policyBase\Profiles", "$policyBase\ODFC", "$policyBase\Logging", "$policyBase\Apps",
    "$directBase\Profiles",  "$directBase\ODFC",  "$directBase\Logging",  "$directBase\Apps"
)
$knownNames = $fslogixDefs.Name | Select-Object -Unique

foreach ($keyPath in $allSubKeys) {
    if (-not (Test-Path $keyPath)) { continue }
    $props = Get-ItemProperty $keyPath -ErrorAction SilentlyContinue
    $props.PSObject.Properties |
        Where-Object { $_.Name -notmatch '^PS' -and $_.Name -notin $knownNames } |
        ForEach-Object {
            $segment = Split-Path $keyPath -Leaf
            $src = if ($keyPath -like '*Policies*') { 'Policy (ADMX/GPO)' } else { 'Direct (Registry)' }
            $extraResults.Add(@{
                Section     = "Additional — $segment"
                Friendly    = $_.Name
                RegPath     = $keyPath
                RegName     = $_.Name
                Value       = $_.Value
                RawValue    = $_.Value
                Source      = $src
                Status      = 'Extra'
                Description = 'Value found in registry not in standard FSLogix ADMX definition list.'
            })
        }
}

Write-Host "      $($results.Count) defined settings read | $($extraResults.Count) additional values found" -ForegroundColor Gray

#endregion

#region --- Also Read Local Group Membership ---

Write-Host "[2/3] Reading FSLogix local security groups..." -ForegroundColor Cyan

$groupInfo = [System.Collections.Generic.List[hashtable]]::new()
$fslogixGroups = @(
    'FSLogix Profile Include List'
    'FSLogix Profile Exclude List'
    'FSLogix ODFC Include List'
    'FSLogix ODFC Exclude List'
)

foreach ($grp in $fslogixGroups) {
    $exists  = $false
    $members = '—'
    try {
        $groupObj = Get-LocalGroup -Name $grp -ErrorAction Stop
        $exists = $true
        $memberList = (Get-LocalGroupMember -Group $grp -ErrorAction SilentlyContinue).Name
        $members = if ($memberList) { $memberList -join ', ' } else { '(empty — no members)' }
    } catch {
        # Group does not exist
    }
    $groupInfo.Add(@{
        Name    = $grp
        Exists  = $exists
        Members = $members
    })
}

#endregion

#region --- Build HTML ---

Write-Host "[3/3] Generating HTML report..." -ForegroundColor Cyan

# Group results by Section
$sections = $results | Group-Object { $_.Section }

$configuredCount    = ($results | Where-Object { $_.Status -ne 'NotConfigured' }).Count
$policyCount        = ($results | Where-Object { $_.Status -eq 'Policy'        }).Count
$directCount        = ($results | Where-Object { $_.Status -eq 'Direct'        }).Count
$notConfiguredCount = ($results | Where-Object { $_.Status -eq 'NotConfigured' }).Count
$extraCount         = $extraResults.Count

# Status badge colors
function Get-StatusBadge($status) {
    switch ($status) {
        'Policy'        { return '<span style="background:#1a6e9e;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;white-space:nowrap">POLICY (GPO)</span>' }
        'Direct'        { return '<span style="background:#5a7a2e;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;white-space:nowrap">DIRECT (REG)</span>' }
        'NotConfigured' { return '<span style="background:#aaa;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;white-space:nowrap">NOT SET</span>' }
        'Extra'         { return '<span style="background:#7b4f9e;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em;white-space:nowrap">EXTRA</span>' }
        default         { return $status }
    }
}

function Get-RowBg($status) {
    switch ($status) {
        'Policy'        { return '#e8f4fd' }
        'Direct'        { return '#edf7e6' }
        'NotConfigured' { return '#f8f9fa' }
        'Extra'         { return '#f5eef8' }
        default         { return '#fff' }
    }
}

# Build section tables
$sectionHtml = ''
foreach ($sec in $sections) {
    $rows = $sec.Group | ForEach-Object {
        $r   = $_
        $bg  = Get-RowBg $r.Status
        $bdg = Get-StatusBadge $r.Status
        $val = [System.Web.HttpUtility]::HtmlEncode($r.Value)
        $reg = [System.Web.HttpUtility]::HtmlEncode("$($r.RegPath)\$($r.RegName)")
        $desc = [System.Web.HttpUtility]::HtmlEncode($r.Description)
        @"
        <tr style="background:$bg">
          <td>$bdg</td>
          <td><strong>$([System.Web.HttpUtility]::HtmlEncode($r.Friendly))</strong><br><span class="desc">$desc</span></td>
          <td class="mono">$val</td>
          <td class="mono small">$reg</td>
        </tr>
"@
    }

    $sectionHtml += @"
    <h2>$([System.Web.HttpUtility]::HtmlEncode($sec.Name))</h2>
    <table>
      <thead><tr>
        <th style="width:120px">Source</th>
        <th>Setting</th>
        <th style="width:220px">Current Value</th>
        <th style="width:300px">Registry Path \ Value Name</th>
      </tr></thead>
      <tbody>$($rows -join '')</tbody>
    </table>
"@
}

# Extra/undocumented settings
if ($extraResults.Count -gt 0) {
    $extraRows = $extraResults | ForEach-Object {
        $r   = $_
        $bg  = Get-RowBg 'Extra'
        $bdg = Get-StatusBadge 'Extra'
        $val = [System.Web.HttpUtility]::HtmlEncode($r.Value)
        $reg = [System.Web.HttpUtility]::HtmlEncode("$($r.RegPath)\$($r.RegName)")
        @"
        <tr style="background:$bg">
          <td>$bdg</td>
          <td><strong>$([System.Web.HttpUtility]::HtmlEncode($r.Friendly))</strong><br><span class="desc">$([System.Web.HttpUtility]::HtmlEncode($r.Description))</span></td>
          <td class="mono">$val</td>
          <td class="mono small">$reg</td>
        </tr>
"@
    }
    $sectionHtml += @"
    <h2>Additional / Undocumented Registry Values</h2>
    <p style="color:#666;font-size:0.9em">These values were found in the FSLogix registry keys but are not in the standard ADMX definition list. They may be custom, legacy, or undocumented settings.</p>
    <table>
      <thead><tr>
        <th style="width:120px">Source</th>
        <th>Value Name</th>
        <th style="width:220px">Current Value</th>
        <th style="width:300px">Registry Path</th>
      </tr></thead>
      <tbody>$($extraRows -join '')</tbody>
    </table>
"@
}

# Local Groups table
$groupRows = $groupInfo | ForEach-Object {
    $g   = $_
    $bg  = if ($g.Exists) { '#e8f4fd' } else { '#fff3cd' }
    $existBadge = if ($g.Exists) {
        '<span style="background:#28a745;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em">EXISTS</span>'
    } else {
        '<span style="background:#dc3545;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.78em">MISSING</span>'
    }
    @"
    <tr style="background:$bg">
      <td>$existBadge</td>
      <td><strong>$([System.Web.HttpUtility]::HtmlEncode($g.Name))</strong></td>
      <td class="mono">$([System.Web.HttpUtility]::HtmlEncode($g.Members))</td>
    </tr>
"@
}

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>FSLogix Settings Audit — $(Get-Date -Format 'yyyy-MM-dd HH:mm')</title>
  <style>
    body   { font-family: Segoe UI, Arial, sans-serif; margin: 30px; background: #f4f6f8; color: #222; }
    h1     { color: #1a3a5c; border-bottom: 3px solid #1a3a5c; padding-bottom: 8px; margin-bottom: 4px; }
    h2     { color: #1a6e9e; margin-top: 36px; margin-bottom: 8px; font-size: 1.1em;
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
    .legend { display:flex; gap:12px; flex-wrap:wrap; margin:8px 0 20px 0; font-size:0.85em; align-items:center; }
  </style>
</head>
<body>

<h1>FSLogix Configuration Settings Audit</h1>

<div class="meta">
  <strong>Machine:</strong> $($env:COMPUTERNAME) &nbsp;|&nbsp;
  <strong>Domain:</strong> $($env:USERDOMAIN) &nbsp;|&nbsp;
  <strong>Run by:</strong> $($env:USERNAME) &nbsp;|&nbsp;
  <strong>Date:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') &nbsp;|&nbsp;
  <strong>OS:</strong> $((Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption)
</div>

<div class="note">
  <strong>Scope:</strong> This report covers
  <code>HKLM:\SOFTWARE\Policies\FSLogix\</code> (GPO/ADMX-managed — takes precedence) and
  <code>HKLM:\SOFTWARE\FSLogix\</code> (direct registry configuration).
  Settings sourced from Local Computer Policy &rarr; Computer Configuration &rarr; Administrative Templates &rarr; FSLogix.
</div>

<h2 style="border:none;margin-top:10px">Summary</h2>
<div class="summary">
  <div class="badge" style="background:#1a6e9e">Policy (GPO)<br>$policyCount</div>
  <div class="badge" style="background:#5a7a2e">Direct (Reg)<br>$directCount</div>
  <div class="badge" style="background:#aaa">Not Set<br>$notConfiguredCount</div>
  <div class="badge" style="background:#7b4f9e">Extra Values<br>$extraCount</div>
</div>

<div class="legend">
  <strong>Legend:</strong>
  <span style="background:#1a6e9e;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.8em">POLICY (GPO)</span> Applied via ADMX/Local GPO — enforced &nbsp;
  <span style="background:#5a7a2e;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.8em">DIRECT (REG)</span> Set directly in registry — not enforced by policy &nbsp;
  <span style="background:#aaa;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.8em">NOT SET</span> Not configured — FSLogix uses default value &nbsp;
  <span style="background:#7b4f9e;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.8em">EXTRA</span> Found in registry, not in standard ADMX list
</div>

$sectionHtml

<h2>FSLogix Local Security Groups</h2>
<table>
  <thead><tr>
    <th style="width:100px">Status</th>
    <th>Group Name</th>
    <th>Members</th>
  </tr></thead>
  <tbody>$($groupRows -join '')</tbody>
</table>

<p style="margin-top:40px;color:#999;font-size:0.8em">
  Generated by Get-FSLogixSettings.ps1 &mdash; PowerShell $($PSVersionTable.PSVersion) &mdash;
  Registry paths: <code>HKLM:\SOFTWARE\Policies\FSLogix</code> &amp; <code>HKLM:\SOFTWARE\FSLogix</code>
</p>

</body>
</html>
"@

$html | Out-File -FilePath $ReportPath -Encoding UTF8
Write-Host "Report saved: $ReportPath" -ForegroundColor Green

try { Start-Process $ReportPath } catch { }

Write-Host "`nDone — Configured: $configuredCount  Policy: $policyCount  Direct: $directCount  Not Set: $notConfiguredCount  Extra: $extraCount`n" -ForegroundColor Cyan
