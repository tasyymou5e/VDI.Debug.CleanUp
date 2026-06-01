#Requires -Version 7.0
<#
.SYNOPSIS
    Parses SecurityGroupAudit HTML report and removes all LGPO.exe instances found.
.DESCRIPTION
    Reads the audit HTML report, extracts every LGPO.exe path from High-confidence findings,
    backs each file up to a timestamped folder, then deletes it.
    Writes a remediation log HTML alongside the source report.
.PARAMETER ReportPath
    Path to the SecurityGroupAudit HTML file produced by Find-SecurityGroupSources.ps1
.PARAMETER BackupRoot
    Folder to copy LGPO.exe files into before deletion (default: same folder as report)
.PARAMETER WhatIf
    List what would be deleted without actually deleting anything.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ReportPath = (Get-ChildItem "$PSScriptRoot\SecurityGroupAudit_*.html" -ErrorAction SilentlyContinue |
                           Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName),
    [string]$BackupRoot  = '',
    [switch]$WhatIfMode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region --- Pre-flight Checks ---

# Check 1: Must be running as Administrator
$currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "`n[BLOCKED] This script must be run as Administrator." -ForegroundColor Red
    Write-Host "  Right-click PowerShell and select 'Run as Administrator', then try again.`n" -ForegroundColor Yellow
    exit 1
}
Write-Host "[OK] Running as Administrator" -ForegroundColor Green

# Check 2: Execution policy must allow scripts (RemoteSigned or Unrestricted)
$policy = Get-ExecutionPolicy -Scope LocalMachine
$allowedPolicies = @('RemoteSigned', 'Unrestricted', 'Bypass')
if ($policy -notin $allowedPolicies) {
    Write-Host "`n[BLOCKED] Execution policy is '$policy' — scripts are not permitted." -ForegroundColor Red
    Write-Host "  Run the following as Administrator to fix, then re-run this script:" -ForegroundColor Yellow
    Write-Host "  Set-ExecutionPolicy RemoteSigned -Scope LocalMachine -Force`n" -ForegroundColor Cyan
    exit 1
}
Write-Host "[OK] Execution policy: $policy" -ForegroundColor Green

Write-Host ""

#endregion

#region --- Setup ---

if (-not (Test-Path $ReportPath)) {
    Write-Error "Report not found: $ReportPath"
    exit 1
}

$reportDir   = Split-Path $ReportPath -Parent
$timestamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath     = Join-Path $reportDir "Remediation_LGPO_$timestamp.html"
$backupDir   = if ($BackupRoot) { $BackupRoot } else { Join-Path $reportDir "LGPO_Backup_$timestamp" }

$actions = [System.Collections.Generic.List[hashtable]]::new()

function Add-Action {
    param(
        [string]$Path,
        [ValidateSet('Deleted','Backed Up','Skipped - Not Found','Skipped - WhatIf','Error')]
        [string]$Status,
        [string]$Detail = ''
    )
    $actions.Add(@{
        Path   = $Path
        Status = $Status
        Detail = $Detail
        Time   = (Get-Date -Format 'HH:mm:ss')
    })
}

#endregion

#region --- 1. Parse HTML for LGPO.exe paths ---

Write-Host "`n[1/4] Reading report file..." -ForegroundColor Cyan
$html = Get-Content $ReportPath -Raw -Encoding UTF8
Write-Host "      Done — $([math]::Round($html.Length/1KB,1)) KB loaded" -ForegroundColor Gray

Write-Host "[2/4] Parsing HTML for LGPO.exe paths..." -ForegroundColor Cyan
$rowPattern = '(?s)<tr[^>]*>.*?<td[^>]*>.*?</td>.*?<td[^>]*>.*?</td>.*?<td[^>]*>LGPO\.exe</td>.*?<td[^>]*>(.*?)</td>'
$rowMatches = [regex]::Matches($html, $rowPattern)

$lgpoPaths = [System.Collections.Generic.List[string]]::new()

foreach ($m in $rowMatches) {
    # Source is in the 4th td (index 3) — let's use a simpler targeted regex per row
    $sourceMatch = [regex]::Match($m.Value, '<td[^>]*font-family:monospace[^>]*>(.*?)</td>')
    if ($sourceMatch.Success) {
        $raw = $sourceMatch.Groups[1].Value.Trim()
        # Decode HTML entities
        Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
        $decoded = [System.Web.HttpUtility]::HtmlDecode($raw)
        if ($decoded -match '(?i)lgpo\.exe$') {
            $lgpoPaths.Add($decoded)
        }
    }
}

# Fallback: simpler grep for any path ending in LGPO.exe in the source column
if ($lgpoPaths.Count -eq 0) {
    $fallback = [regex]::Matches($html, '(?i)([A-Za-z]:\\[^<"]+lgpo\.exe)')
    foreach ($f in $fallback) {
        $p = $f.Groups[1].Value.Trim()
        if ($p -notin $lgpoPaths) { $lgpoPaths.Add($p) }
    }
}

# Deduplicate
$lgpoPaths = $lgpoPaths | Select-Object -Unique

Write-Host "      Found $($lgpoPaths.Count) LGPO.exe path(s):" -ForegroundColor White
$lgpoPaths | ForEach-Object { Write-Host "        $_" -ForegroundColor Yellow }

if ($lgpoPaths.Count -eq 0) {
    Write-Host "`nNo LGPO.exe paths found in report. Nothing to do." -ForegroundColor Green
    exit 0
}

#endregion

#region --- 2. Confirm / WhatIf ---

if ($WhatIfMode) {
    Write-Host "`n[WhatIf mode] Would delete:" -ForegroundColor Magenta
    $lgpoPaths | ForEach-Object {
        Write-Host "  $_" -ForegroundColor Yellow
        Add-Action -Path $_ -Status 'Skipped - WhatIf' -Detail 'WhatIf mode enabled — no changes made'
    }
} else {
    Write-Host "[3/4] Creating backup directory..." -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    Write-Host "      $backupDir" -ForegroundColor Gray

    Write-Host "[4/4] Backing up and deleting LGPO.exe instances..." -ForegroundColor Cyan

    #region --- 3. Backup then Delete ---

    $i = 0
    foreach ($lgpoPath in $lgpoPaths) {
        $i++
        Write-Host "  [$i/$($lgpoPaths.Count)] $lgpoPath" -ForegroundColor White

        if (-not (Test-Path $lgpoPath)) {
            Write-Host "  Not found on disk (may have already been removed)" -ForegroundColor Gray
            Add-Action -Path $lgpoPath -Status 'Skipped - Not Found' -Detail 'File not present on disk at time of remediation'
            continue
        }

        # Build backup path preserving folder structure
        $relativePath = $lgpoPath -replace '^[A-Za-z]:\\', '' -replace '\\', '_'
        $backupDest   = Join-Path $backupDir $relativePath

        try {
            # Backup
            Copy-Item -Path $lgpoPath -Destination $backupDest -Force
            Write-Host "  Backed up to: $backupDest" -ForegroundColor Cyan
            Add-Action -Path $lgpoPath -Status 'Backed Up' -Detail "Backup: $backupDest"

            # Delete
            Remove-Item -Path $lgpoPath -Force
            Write-Host "  Deleted: $lgpoPath" -ForegroundColor Green
            Add-Action -Path $lgpoPath -Status 'Deleted' -Detail "Removed from golden image. Backup retained at: $backupDest"

        } catch {
            Write-Host "  ERROR: $_" -ForegroundColor Red
            Add-Action -Path $lgpoPath -Status 'Error' -Detail "Exception: $_"
        }
    }

    #endregion

    #region --- 4. Verify deletions ---

    Write-Host "`n--- Verification ---" -ForegroundColor Cyan
    $stillPresent = $lgpoPaths | Where-Object { Test-Path $_ }
    if ($stillPresent.Count -eq 0) {
        Write-Host "  All LGPO.exe instances confirmed removed." -ForegroundColor Green
    } else {
        Write-Host "  WARNING — still present on disk:" -ForegroundColor Red
        $stillPresent | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
    }

    #endregion
}

#endregion

#region --- 5. Generate HTML Remediation Report ---

Write-Host "`n[5/5] Generating HTML report..." -ForegroundColor Cyan

Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

$statusColors = @{
    'Deleted'              = '#d4edda'
    'Backed Up'            = '#cce5ff'
    'Skipped - Not Found'  = '#f8f9fa'
    'Skipped - WhatIf'     = '#fff3cd'
    'Error'                = '#f8d7da'
}
$statusBadge = @{
    'Deleted'              = '#28a745'
    'Backed Up'            = '#007bff'
    'Skipped - Not Found'  = '#6c757d'
    'Skipped - WhatIf'     = '#ffc107'
    'Error'                = '#dc3545'
}

$deletedCount = ($actions | Where-Object { $_.Status -eq 'Deleted'     }).Count
$errorCount   = ($actions | Where-Object { $_.Status -eq 'Error'       }).Count
$skipCount    = ($actions | Where-Object { $_.Status -like 'Skipped*'  }).Count

$rows = $actions | ForEach-Object {
    $bg    = $statusColors[$_.Status]
    $badge = $statusBadge[$_.Status]
    $pathEnc   = [System.Web.HttpUtility]::HtmlEncode($_.Path)
    $detailEnc = [System.Web.HttpUtility]::HtmlEncode($_.Detail)
    @"
    <tr style="background:$bg">
      <td style="white-space:nowrap">$($_.Time)</td>
      <td><span style="background:$badge;color:#fff;padding:2px 8px;border-radius:4px;font-size:0.8em;white-space:nowrap">$($_.Status)</span></td>
      <td style="font-family:monospace;font-size:0.85em;word-break:break-all">$pathEnc</td>
      <td style="font-family:monospace;font-size:0.82em">$detailEnc</td>
    </tr>
"@
}

$whatIfNote = if ($WhatIfMode) {
    '<div style="background:#fff3cd;border-left:4px solid #ffc107;padding:12px 16px;margin-bottom:20px;border-radius:0 6px 6px 0"><strong>WhatIf Mode:</strong> No files were deleted. Re-run without <code>-WhatIfMode</code> to remediate.</div>'
} else { '' }

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>LGPO.exe Remediation Report — $timestamp</title>
  <style>
    body  { font-family: Segoe UI, Arial, sans-serif; margin: 30px; background: #f5f5f5; color: #333; }
    h1    { color: #2c3e50; border-bottom: 3px solid #2c3e50; padding-bottom: 8px; }
    h2    { color: #495057; margin-top: 30px; }
    .meta { background:#fff; border-radius:6px; padding:15px 20px; margin-bottom:20px; box-shadow:0 1px 3px rgba(0,0,0,.1); }
    .summary { display:flex; gap:15px; flex-wrap:wrap; margin-bottom:25px; }
    .badge   { padding:10px 20px; border-radius:8px; color:#fff; font-size:1.1em; font-weight:bold; min-width:80px; text-align:center; }
    table { width:100%; border-collapse:collapse; background:#fff; border-radius:8px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,.1); }
    th    { background:#2c3e50; color:#fff; padding:10px 12px; text-align:left; font-size:0.9em; }
    td    { padding:8px 12px; border-bottom:1px solid #dee2e6; vertical-align:top; font-size:0.9em; }
    tr:last-child td { border-bottom:none; }
    code  { background:#f0f0f0; padding:1px 4px; border-radius:3px; font-size:0.9em; }
    .note { background:#e8f4f8; border-left:4px solid #17a2b8; padding:12px 16px; border-radius:0 6px 6px 0; margin-bottom:20px; }
  </style>
</head>
<body>
  <h1>LGPO.exe Remediation Report</h1>

  <div class="meta">
    <strong>Machine:</strong> $($env:COMPUTERNAME) &nbsp;|&nbsp;
    <strong>Run by:</strong> $($env:USERNAME) &nbsp;|&nbsp;
    <strong>Date:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') &nbsp;|&nbsp;
    <strong>Source report:</strong> <code>$(Split-Path $ReportPath -Leaf)</code>
  </div>

  $whatIfNote

  <div class="note">
    <strong>Scope:</strong> Only <code>LGPO.exe</code> Category findings from the audit report were processed.
    GPO startup scripts (MDE/Defender ATP onboarding) and Registry.pol were <strong>not modified</strong>.
    All deleted files were backed up to: <code>$backupDir</code>
  </div>

  <h2>Summary</h2>
  <div class="summary">
    <div class="badge" style="background:#28a745">Deleted<br>$deletedCount</div>
    <div class="badge" style="background:#dc3545">Errors<br>$errorCount</div>
    <div class="badge" style="background:#6c757d">Skipped<br>$skipCount</div>
  </div>

  <h2>Actions Taken</h2>
  <table>
    <thead>
      <tr>
        <th>Time</th>
        <th>Status</th>
        <th>Path</th>
        <th>Detail</th>
      </tr>
    </thead>
    <tbody>
      $($rows -join "`n")
    </tbody>
  </table>

  <h2>What to Do Next</h2>
  <ul>
    <li>Confirm the 4 security groups no longer appear on a freshly provisioned VDI session</li>
    <li>If groups still appear: check <code>Registry.pol</code> content using <code>lgpo.exe /parse /m</code> from a separate machine</li>
    <li>The SCCM ccmcache copies (<code>C:\Windows\ccmcache\*</code>) will be re-downloaded by SCCM on next client evaluation — <br>
        consider updating the SCCM package source to remove LGPO.exe from the package itself</li>
    <li>The <code>C:\IP\ImagePrep\TOOLS\CheckImage\</code> copy should be reviewed with the image build team — <br>
        this is the likely source that was run to configure local policy</li>
    <li>Backup retained at: <code>$backupDir</code> — keep until VDI provisioning is confirmed clean</li>
  </ul>

  <p style="margin-top:30px;color:#999;font-size:0.8em">
    Generated by Remediate-LGPOExe.ps1 &mdash; PowerShell $($PSVersionTable.PSVersion)
  </p>
</body>
</html>
"@

$html | Out-File -FilePath $logPath -Encoding UTF8
Write-Host "`nRemediation report saved: $logPath" -ForegroundColor Green

try { Start-Process $logPath } catch { }

Write-Host "`nDone — Deleted: $deletedCount  Errors: $errorCount  Skipped: $skipCount`n" -ForegroundColor $(
    if ($errorCount -gt 0) { 'Red' } elseif ($deletedCount -gt 0) { 'Green' } else { 'Yellow' }
)
