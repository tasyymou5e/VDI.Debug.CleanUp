#Requires -Version 7.0
<#
.SYNOPSIS
    Audits all potential sources that create local security groups on a VDI golden image.
.DESCRIPTION
    Checks GPO settings, LGPO, OSOT, Horizon provisioning scripts, startup scripts,
    scheduled tasks, and registry run keys. Outputs a timestamped HTML report.
.NOTES
    Run as Administrator on the golden image or an active VDI session.
#>

[CmdletBinding()]
param(
    [string]$ReportPath = "C:\Temp\SecurityGroupAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').html",
    [string[]]$GroupNamesToSearch = @()   # optional: filter findings to specific group names
)

#region --- Helpers ---

$findings = [System.Collections.Generic.List[hashtable]]::new()

function Add-Finding {
    param(
        [string]$Category,
        [string]$Source,
        [string]$Detail,
        [ValidateSet('High','Medium','Low','Info')]
        [string]$Confidence = 'Medium'
    )
    $findings.Add(@{
        Category   = $Category
        Source     = $Source
        Detail     = $Detail
        Confidence = $Confidence
        Time       = (Get-Date -Format 'HH:mm:ss')
    })
}

function Write-Section ([string]$Name) {
    Write-Host "`n=== $Name ===" -ForegroundColor Cyan
}

function Test-ContainsGroupName ([string]$Content) {
    if ($GroupNamesToSearch.Count -eq 0) { return $true }
    foreach ($name in $GroupNamesToSearch) {
        if ($Content -match [regex]::Escape($name)) { return $true }
    }
    return $false
}

#endregion

#region --- 1. Enumerate Current Local Groups ---

Write-Section "1. Current Local Security Groups"

try {
    $localGroups = Get-LocalGroup -ErrorAction Stop
    foreach ($g in $localGroups) {
        $members = (Get-LocalGroupMember -Group $g.Name -ErrorAction SilentlyContinue).Name -join ', '
        Add-Finding -Category 'Local Groups' `
                    -Source 'Get-LocalGroup' `
                    -Detail "Group: $($g.Name) | SID: $($g.SID) | Members: $(if($members){"$members"}else{'(none)'})" `
                    -Confidence 'Info'
        Write-Host "  $($g.Name)  [$($g.SID)]" -ForegroundColor White
    }
} catch {
    Add-Finding -Category 'Local Groups' -Source 'Get-LocalGroup' -Detail "ERROR: $_" -Confidence 'Info'
}

#endregion

#region --- 2. Local GPO Security Template (LGPO / Restricted Groups) ---

Write-Section "2. Local GPO Security Template (GptTmpl.inf)"

$gptTmpl = "C:\Windows\System32\GroupPolicy\Machine\Microsoft\Windows NT\SecEdit\GptTmpl.inf"

if (Test-Path $gptTmpl) {
    $content = Get-Content $gptTmpl -Raw -ErrorAction SilentlyContinue

    # Group Membership section = Restricted Groups
    if ($content -match '\[Group Membership\]') {
        $section = [regex]::Match($content, '(?s)\[Group Membership\](.*?)(\[|$)').Groups[1].Value
        if (Test-ContainsGroupName $section) {
            Add-Finding -Category 'LGPO / Restricted Groups' `
                        -Source $gptTmpl `
                        -Detail "Found [Group Membership] section:`n$($section.Trim())" `
                        -Confidence 'High'
            Write-Host "  [HIGH] Restricted Groups / Group Membership entries found in GptTmpl.inf" -ForegroundColor Red
        }
    } else {
        Write-Host "  No [Group Membership] section found in GptTmpl.inf" -ForegroundColor Gray
    }

    # Show full file as Info
    Add-Finding -Category 'LGPO / Restricted Groups' `
                -Source $gptTmpl `
                -Detail "Full GptTmpl.inf content:`n$content" `
                -Confidence 'Info'
} else {
    Add-Finding -Category 'LGPO / Restricted Groups' `
                -Source $gptTmpl `
                -Detail 'File not found — no local GPO security template applied' `
                -Confidence 'Info'
    Write-Host "  GptTmpl.inf not found" -ForegroundColor Gray
}

#endregion

#region --- 3. Local GPO Registry.pol (Machine) ---

Write-Section "3. Local GPO Registry.pol"

$regPol = "C:\Windows\System32\GroupPolicy\Machine\Registry.pol"
if (Test-Path $regPol) {
    $size = (Get-Item $regPol).Length
    Add-Finding -Category 'Local GPO Registry.pol' `
                -Source $regPol `
                -Detail "File exists — size: $size bytes. Use 'lgpo.exe /parse /m `"$regPol`"' to decode." `
                -Confidence 'Medium'
    Write-Host "  Registry.pol exists ($size bytes) — may contain group-related policy" -ForegroundColor Yellow

    # Try to parse with lgpo.exe if present
    $lgpoExe = Get-ChildItem C:\, 'C:\Windows\System32\', 'C:\Windows\SysWOW64\', 'C:\Temp\', 'C:\Tools\' `
                    -Filter lgpo.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty FullName
    if ($lgpoExe) {
        $parsed = & $lgpoExe /parse /m $regPol 2>&1
        Add-Finding -Category 'Local GPO Registry.pol' `
                    -Source "lgpo.exe parse" `
                    -Detail ($parsed -join "`n") `
                    -Confidence 'High'
        Write-Host "  Parsed with lgpo.exe" -ForegroundColor Green
    }
} else {
    Write-Host "  Registry.pol not found" -ForegroundColor Gray
}

#endregion

#region --- 4. LGPO.exe Presence ---

Write-Section "4. LGPO.exe"

$lgpoPaths = @(
    'C:\Windows\System32\lgpo.exe'
    'C:\Windows\SysWOW64\lgpo.exe'
    'C:\Temp\lgpo.exe'
    'C:\Tools\lgpo.exe'
    'C:\Windows\Scripts\lgpo.exe'
    'C:\Windows\Setup\Scripts\lgpo.exe'
)

$lgpoFound = $false
foreach ($p in $lgpoPaths) {
    if (Test-Path $p) {
        $ver = (Get-Item $p).VersionInfo.FileVersion
        Add-Finding -Category 'LGPO.exe' `
                    -Source $p `
                    -Detail "lgpo.exe found — version: $ver. Indicates local policy was imported via LGPO." `
                    -Confidence 'High'
        Write-Host "  [HIGH] lgpo.exe found at: $p  (v$ver)" -ForegroundColor Red
        $lgpoFound = $true
    }
}

# Also do a broader search
$lgpoBroad = Get-ChildItem C:\ -Filter lgpo.exe -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $_.FullName -notin $lgpoPaths }
foreach ($f in $lgpoBroad) {
    Add-Finding -Category 'LGPO.exe' `
                -Source $f.FullName `
                -Detail "lgpo.exe found (broad search) — version: $($f.VersionInfo.FileVersion)" `
                -Confidence 'High'
    Write-Host "  [HIGH] lgpo.exe found at: $($f.FullName)" -ForegroundColor Red
    $lgpoFound = $true
}

if (-not $lgpoFound) { Write-Host "  lgpo.exe not found on this machine" -ForegroundColor Gray }

#endregion

#region --- 5. GPO Startup Scripts ---

Write-Section "5. GPO Startup Scripts"

$startupScriptDir = "C:\Windows\System32\GroupPolicy\Machine\Scripts\Startup"
if (Test-Path $startupScriptDir) {
    $scripts = Get-ChildItem $startupScriptDir -Recurse -ErrorAction SilentlyContinue
    foreach ($s in $scripts) {
        $content = if ($s.Extension -in '.ps1','.bat','.cmd','.vbs') {
            Get-Content $s.FullName -Raw -ErrorAction SilentlyContinue
        } else { '(binary/non-text)' }

        $confidence = if (Test-ContainsGroupName ($content ?? '')) { 'High' } else { 'Medium' }
        Add-Finding -Category 'GPO Startup Script' `
                    -Source $s.FullName `
                    -Detail "Script found:`n$content" `
                    -Confidence $confidence
        Write-Host "  Script: $($s.FullName)" -ForegroundColor Yellow
    }
} else {
    Write-Host "  No GPO startup scripts directory found" -ForegroundColor Gray
}

# Check scripts.ini for registered startup scripts
$scriptsIni = "C:\Windows\System32\GroupPolicy\Machine\Scripts\scripts.ini"
if (Test-Path $scriptsIni) {
    $iniContent = Get-Content $scriptsIni -Raw
    Add-Finding -Category 'GPO Startup Script' `
                -Source $scriptsIni `
                -Detail "scripts.ini content:`n$iniContent" `
                -Confidence 'High'
    Write-Host "  [HIGH] scripts.ini found with content" -ForegroundColor Red
}

#endregion

#region --- 6. Common Script Locations ---

Write-Section "6. Common Script / Setup Locations"

$scriptPaths = @(
    'C:\Windows\Scripts'
    'C:\Windows\Setup\Scripts'
    'C:\Windows\Panther'
    'C:\Windows\Temp'
    'C:\Temp'
    'C:\ProgramData\VMware'
    'C:\ProgramData\Omnissa'
    'C:\Program Files\VMware'
)

$scriptExtensions = @('.ps1', '.bat', '.cmd', '.vbs', '.wsf')

foreach ($dir in $scriptPaths) {
    if (-not (Test-Path $dir)) { continue }
    $files = Get-ChildItem $dir -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -in $scriptExtensions }
    foreach ($f in $files) {
        $content = Get-Content $f.FullName -Raw -ErrorAction SilentlyContinue
        $mentionsGroups = $content -match 'localgroup|LocalGroup|net group|Add-LocalGroup|New-LocalGroup|security group'
        $confidence = if ($mentionsGroups -and (Test-ContainsGroupName ($content ?? ''))) { 'High' }
                      elseif ($mentionsGroups) { 'Medium' }
                      else { 'Low' }

        if ($mentionsGroups) {
            Add-Finding -Category 'Script File' `
                        -Source $f.FullName `
                        -Detail "References local group operations:`n$(($content -split "`n" | Where-Object { $_ -match 'localgroup|LocalGroup|net group|Add-LocalGroup|New-LocalGroup' }) -join "`n")" `
                        -Confidence $confidence
            Write-Host "  [$confidence] Group-related script: $($f.FullName)" -ForegroundColor $(if ($confidence -eq 'High') { 'Red' } else { 'Yellow' })
        }
    }
}

#endregion

#region --- 7. OSOT Detection ---

Write-Section "7. VMware OS Optimization Tool (OSOT)"

$osotPaths = @(
    'C:\Program Files\VMware\VMware OS Optimization Tool'
    'C:\Program Files (x86)\VMware\VMware OS Optimization Tool'
    'C:\ProgramData\VMware\VMware OS Optimization Tool'
)

foreach ($p in $osotPaths) {
    if (Test-Path $p) {
        $files = Get-ChildItem $p -Recurse -ErrorAction SilentlyContinue
        Add-Finding -Category 'OSOT' `
                    -Source $p `
                    -Detail "OSOT installation found. Files:`n$(($files.FullName) -join "`n")" `
                    -Confidence 'High'
        Write-Host "  [HIGH] OSOT found at: $p" -ForegroundColor Red

        # Look for XML templates that reference groups
        $templates = $files | Where-Object { $_.Extension -eq '.xml' }
        foreach ($t in $templates) {
            $xml = Get-Content $t.FullName -Raw -ErrorAction SilentlyContinue
            if ($xml -match 'LocalGroup|localgroup|security group|Group') {
                Add-Finding -Category 'OSOT Template' `
                            -Source $t.FullName `
                            -Detail "OSOT template references groups:`n$(($xml -split "`n" | Where-Object { $_ -match 'Group' }) -join "`n")" `
                            -Confidence 'High'
                Write-Host "    Template with group refs: $($t.FullName)" -ForegroundColor Red
            }
        }
    }
}

# OSOT log files
$logLocations = @('C:\Windows\Logs', 'C:\Windows\Temp', 'C:\ProgramData\VMware')
foreach ($loc in $logLocations) {
    if (-not (Test-Path $loc)) { continue }
    Get-ChildItem $loc -Filter '*osot*' -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $content = Get-Content $_.FullName -Raw -ErrorAction SilentlyContinue
        Add-Finding -Category 'OSOT Log' `
                    -Source $_.FullName `
                    -Detail "OSOT log found:`n$content" `
                    -Confidence 'High'
        Write-Host "  OSOT log: $($_.FullName)" -ForegroundColor Yellow
    }
}

# Event log
try {
    $osotEvents = Get-WinEvent -LogName Application -ErrorAction Stop |
                    Where-Object { $_.ProviderName -like '*OSOT*' -or $_.Message -like '*OS Optimization*' } |
                    Select-Object -First 20
    foreach ($e in $osotEvents) {
        Add-Finding -Category 'OSOT Event Log' `
                    -Source "Application EventLog @ $($e.TimeCreated)" `
                    -Detail $e.Message `
                    -Confidence 'Medium'
    }
    if ($osotEvents.Count -gt 0) {
        Write-Host "  OSOT events found in Application log: $($osotEvents.Count)" -ForegroundColor Yellow
    }
} catch { }

#endregion

#region --- 8. Horizon / VMware Customization ---

Write-Section "8. Horizon Instant Clone / QuickPrep Scripts"

$horizonScriptPaths = @(
    'C:\Program Files\VMware\VMware Tools\scripts'
    'C:\ProgramData\VMware\VMware Tools'
    'C:\Windows\Temp\vmware-imc'
    'C:\Windows\Setup\Scripts'
)

foreach ($p in $horizonScriptPaths) {
    if (-not (Test-Path $p)) { continue }
    Get-ChildItem $p -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -in $scriptExtensions } |
        ForEach-Object {
            $content = Get-Content $_.FullName -Raw -ErrorAction SilentlyContinue
            $mentionsGroups = $content -match 'localgroup|LocalGroup|net group|Add-LocalGroup|New-LocalGroup'
            $confidence = if ($mentionsGroups) { 'High' } else { 'Low' }
            Add-Finding -Category 'Horizon Provisioning Script' `
                        -Source $_.FullName `
                        -Detail $(if ($mentionsGroups) {
                            "References local group operations:`n$(($content -split "`n" | Where-Object { $_ -match 'localgroup|LocalGroup|net group|Add-LocalGroup|New-LocalGroup' }) -join "`n")"
                        } else { "Script exists (no group refs detected)" }) `
                        -Confidence $confidence
            if ($mentionsGroups) {
                Write-Host "  [HIGH] Group-related Horizon script: $($_.FullName)" -ForegroundColor Red
            }
        }
}

# VMware Tools registry entries
$vmwareRegKeys = @(
    'HKLM:\SOFTWARE\VMware, Inc.\VMware Tools'
    'HKLM:\SOFTWARE\VMware, Inc.\Horizon Agent'
    'HKLM:\SOFTWARE\Omnissa'
)
foreach ($key in $vmwareRegKeys) {
    if (Test-Path $key) {
        $props = Get-ItemProperty $key -ErrorAction SilentlyContinue
        Add-Finding -Category 'VMware/Horizon Registry' `
                    -Source $key `
                    -Detail ($props | Out-String) `
                    -Confidence 'Info'
    }
}

#endregion

#region --- 9. Scheduled Tasks ---

Write-Section "9. Scheduled Tasks (non-Microsoft)"

$tasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
            Where-Object { $_.TaskPath -notlike '\Microsoft\*' }

foreach ($task in $tasks) {
    $actions = $task.Actions | ForEach-Object {
        $cmd = "$($_.Execute) $($_.Arguments)"
        $scriptContent = ''
        # If action calls a script file, read it
        if ($_.Execute -match '\.(ps1|bat|cmd|vbs)$' -and (Test-Path $_.Execute)) {
            $scriptContent = Get-Content $_.Execute -Raw -ErrorAction SilentlyContinue
        }
        "$cmd`n$scriptContent"
    }
    $fullDetail = $actions -join "`n"
    $mentionsGroups = $fullDetail -match 'localgroup|LocalGroup|net group|Add-LocalGroup|New-LocalGroup'
    $confidence = if ($mentionsGroups) { 'High' } else { 'Low' }

    if ($mentionsGroups) {
        Add-Finding -Category 'Scheduled Task' `
                    -Source "$($task.TaskPath)$($task.TaskName)" `
                    -Detail "Task references local group operations:`n$fullDetail" `
                    -Confidence $confidence
        Write-Host "  [HIGH] Group-related scheduled task: $($task.TaskName)" -ForegroundColor Red
    } else {
        Add-Finding -Category 'Scheduled Task' `
                    -Source "$($task.TaskPath)$($task.TaskName)" `
                    -Detail "Actions: $($task.Actions.Execute -join ', ')" `
                    -Confidence 'Info'
        Write-Host "  Task: $($task.TaskName)" -ForegroundColor Gray
    }
}

#endregion

#region --- 10. Registry Run Keys ---

Write-Section "10. Registry Run / RunOnce Keys"

$runKeys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce'
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx'
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'
    'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\BootExecute'
    'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
)

foreach ($key in $runKeys) {
    if (-not (Test-Path $key)) { continue }
    $values = Get-ItemProperty $key -ErrorAction SilentlyContinue
    $values.PSObject.Properties |
        Where-Object { $_.Name -notmatch '^PS' } |
        ForEach-Object {
            $val = $_.Value.ToString()
            $mentionsGroups = $val -match 'localgroup|LocalGroup|net group|lgpo'
            $confidence = if ($mentionsGroups) { 'High' } else { 'Info' }
            Add-Finding -Category 'Registry Run Key' `
                        -Source "$key\$($_.Name)" `
                        -Detail $val `
                        -Confidence $confidence
            if ($mentionsGroups) {
                Write-Host "  [HIGH] Group-related run key: $key\$($_.Name) = $val" -ForegroundColor Red
            }
        }
}

#endregion

#region --- 11. SetupComplete / Unattend ---

Write-Section "11. Sysprep / SetupComplete / Unattend"

$sysprepFiles = @(
    'C:\Windows\Setup\Scripts\SetupComplete.cmd'
    'C:\Windows\Setup\Scripts\SetupComplete.ps1'
    'C:\Windows\Panther\unattend.xml'
    'C:\Windows\System32\sysprep\unattend.xml'
    'C:\Unattend.xml'
)

foreach ($f in $sysprepFiles) {
    if (-not (Test-Path $f)) { continue }
    $content = Get-Content $f -Raw -ErrorAction SilentlyContinue
    $mentionsGroups = $content -match 'localgroup|LocalGroup|net group|Add-LocalGroup|New-LocalGroup|Group'
    $confidence = if ($mentionsGroups) { 'High' } else { 'Medium' }
    Add-Finding -Category 'Sysprep / Setup' `
                -Source $f `
                -Detail "Content:`n$content" `
                -Confidence $confidence
    Write-Host "  [$confidence] Found: $f" -ForegroundColor $(if ($mentionsGroups) { 'Red' } else { 'Yellow' })
}

#endregion

#region --- 12. Event Log — Group Creation Events ---

Write-Section "12. Event Log — Security Group Creation (Event 4731)"

try {
    # 4731 = local security group created
    $groupEvents = Get-WinEvent -FilterHashtable @{
        LogName   = 'Security'
        Id        = @(4731, 4734, 4728, 4732)   # created, deleted, member added, member added
        StartTime = (Get-Date).AddDays(-30)
    } -ErrorAction Stop | Select-Object -First 50

    foreach ($e in $groupEvents) {
        Add-Finding -Category 'Event Log' `
                    -Source "Security Log — EventID $($e.Id) @ $($e.TimeCreated)" `
                    -Detail $e.Message `
                    -Confidence 'High'
    }
    Write-Host "  Group-related security events (last 30d): $($groupEvents.Count)" -ForegroundColor $(if ($groupEvents.Count -gt 0) { 'Yellow' } else { 'Gray' })
} catch {
    Add-Finding -Category 'Event Log' -Source 'Security Log' -Detail "Could not read Security log: $_" -Confidence 'Info'
    Write-Host "  Could not read Security event log (may need admin rights or audit not enabled)" -ForegroundColor Gray
}

#endregion

#region --- Generate HTML Report ---

Write-Section "Generating HTML Report"

$confidenceColors = @{
    High   = '#f8d7da'
    Medium = '#fff3cd'
    Low    = '#d1ecf1'
    Info   = '#f8f9fa'
}
$confidenceBadge = @{
    High   = '#dc3545'
    Medium = '#ffc107'
    Low    = '#17a2b8'
    Info   = '#6c757d'
}

$highCount   = ($findings | Where-Object { $_.Confidence -eq 'High'   }).Count
$medCount    = ($findings | Where-Object { $_.Confidence -eq 'Medium' }).Count
$lowCount    = ($findings | Where-Object { $_.Confidence -eq 'Low'    }).Count
$infoCount   = ($findings | Where-Object { $_.Confidence -eq 'Info'   }).Count

$tableRows = $findings | ForEach-Object {
    $bg    = $confidenceColors[$_.Confidence]
    $badge = $confidenceBadge[$_.Confidence]
    $detailEscaped = [System.Web.HttpUtility]::HtmlEncode($_.Detail) -replace "`n", '<br>'
    @"
    <tr style="background:$bg">
      <td style="white-space:nowrap">$($_.Time)</td>
      <td><span style="background:$badge;color:#fff;padding:2px 7px;border-radius:4px;font-size:0.8em">$($_.Confidence)</span></td>
      <td>$($_.Category)</td>
      <td style="font-family:monospace;font-size:0.85em">$([System.Web.HttpUtility]::HtmlEncode($_.Source))</td>
      <td style="font-family:monospace;font-size:0.82em;white-space:pre-wrap;max-width:600px">$detailEscaped</td>
    </tr>
"@
}

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Security Group Source Audit — $(Get-Date -Format 'yyyy-MM-dd HH:mm')</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 30px; background: #f5f5f5; color: #333; }
    h1   { color: #2c3e50; border-bottom: 3px solid #2c3e50; padding-bottom: 8px; }
    h2   { color: #495057; margin-top: 30px; }
    .meta { background: #fff; border-radius: 6px; padding: 15px 20px; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,.1); }
    .summary { display: flex; gap: 15px; flex-wrap: wrap; margin-bottom: 25px; }
    .badge { padding: 10px 20px; border-radius: 8px; color: #fff; font-size: 1.1em; font-weight: bold; min-width: 80px; text-align: center; }
    .badge-high   { background: #dc3545; }
    .badge-medium { background: #ffc107; color: #333; }
    .badge-low    { background: #17a2b8; }
    .badge-info   { background: #6c757d; }
    table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,.1); }
    th { background: #2c3e50; color: #fff; padding: 10px 12px; text-align: left; font-size: 0.9em; }
    td { padding: 8px 12px; border-bottom: 1px solid #dee2e6; vertical-align: top; font-size: 0.9em; }
    tr:last-child td { border-bottom: none; }
    .note { background: #e8f4f8; border-left: 4px solid #17a2b8; padding: 12px 16px; border-radius: 0 6px 6px 0; margin-bottom: 20px; }
  </style>
</head>
<body>
  <h1>Security Group Source Audit</h1>

  <div class="meta">
    <strong>Machine:</strong> $($env:COMPUTERNAME) &nbsp;|&nbsp;
    <strong>Domain:</strong> $($env:USERDOMAIN) &nbsp;|&nbsp;
    <strong>Run by:</strong> $($env:USERNAME) &nbsp;|&nbsp;
    <strong>Date:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') &nbsp;|&nbsp;
    <strong>OS:</strong> $((Get-CimInstance Win32_OperatingSystem).Caption)
  </div>

  <div class="note">
    <strong>Purpose:</strong> Identifies all potential sources that create local security groups on this VDI golden image —
    including GPO Restricted Groups, LGPO.exe, OSOT templates, Horizon provisioning scripts, startup scripts,
    scheduled tasks, and registry run keys.
  </div>

  <h2>Summary</h2>
  <div class="summary">
    <div class="badge badge-high">High<br>$highCount</div>
    <div class="badge badge-medium">Medium<br>$medCount</div>
    <div class="badge badge-low">Low<br>$lowCount</div>
    <div class="badge badge-info">Info<br>$infoCount</div>
  </div>

  <h2>Findings</h2>
  <table>
    <thead>
      <tr>
        <th>Time</th>
        <th>Confidence</th>
        <th>Category</th>
        <th>Source</th>
        <th>Detail</th>
      </tr>
    </thead>
    <tbody>
      $($tableRows -join "`n")
    </tbody>
  </table>

  <p style="margin-top:30px;color:#999;font-size:0.8em">
    Generated by Find-SecurityGroupSources.ps1 &mdash; PowerShell $($PSVersionTable.PSVersion)
  </p>
</body>
</html>
"@

# HttpUtility may not be loaded in PS7 on Server — load it
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

# Re-render with encoding now loaded
$tableRows = $findings | ForEach-Object {
    $bg    = $confidenceColors[$_.Confidence]
    $badge = $confidenceBadge[$_.Confidence]
    $detailEscaped = [System.Web.HttpUtility]::HtmlEncode($_.Detail) -replace "`n", '<br>'
    @"
    <tr style="background:$bg">
      <td style="white-space:nowrap">$($_.Time)</td>
      <td><span style="background:$badge;color:#fff;padding:2px 7px;border-radius:4px;font-size:0.8em">$($_.Confidence)</span></td>
      <td>$($_.Category)</td>
      <td style="font-family:monospace;font-size:0.85em">$([System.Web.HttpUtility]::HtmlEncode($_.Source))</td>
      <td style="font-family:monospace;font-size:0.82em;white-space:pre-wrap;max-width:600px">$detailEscaped</td>
    </tr>
"@
}

$html = $html -replace '(?s)<tbody>.*</tbody>', "<tbody>`n$($tableRows -join "`n")`n    </tbody>"

$html | Out-File -FilePath $ReportPath -Encoding UTF8
Write-Host "`nReport saved: $ReportPath" -ForegroundColor Green

# Try to open it
try { Start-Process $ReportPath } catch { }

Write-Host "`nDone. High-confidence findings: $highCount" -ForegroundColor $(if ($highCount -gt 0) { 'Red' } else { 'Green' })
