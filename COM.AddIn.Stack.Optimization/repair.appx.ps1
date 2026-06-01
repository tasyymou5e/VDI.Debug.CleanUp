#Requires -Version 7.0
<#
.SYNOPSIS
    Repair / upgrade Windows in-place using Windows Setup (setup.exe).

.DESCRIPTION
    Performs a quiet in-place upgrade/repair of Windows using the Windows Setup
    executable. This is typically used to repair a corrupted Windows installation
    or apply a feature update without wiping the image.

    IMPORTANT — READ BEFORE RUNNING:
      • This script MUST be run from an elevated PowerShell session (Administrator).
      • You MUST supply a valid Windows 11 ISO or extracted setup directory via -SetupPath.
      • The /NoReboot flag is set by default — reboot manually after reviewing the log.
      • This is NOT for golden image sealing. Use it to repair a broken install only.
      • Running setup.exe /auto upgrade on a fully-configured image WILL reset certain
        settings. Re-run all Build scripts after an in-place upgrade.

.PARAMETER SetupPath
    Full path to setup.exe — either from a mounted ISO or an extracted Windows media folder.
    Example: D:\setup.exe  or  C:\WinMedia\setup.exe

.PARAMETER NoReboot
    If specified, suppresses automatic reboot after setup completes. Default: $true.

.PARAMETER DryRun
    Displays the full command that would run without executing it. Use to validate parameters.

.EXAMPLE
    # Verify command before running
    .\repair.appx.ps1 -SetupPath "D:\setup.exe" -DryRun

    # Perform in-place upgrade/repair, no auto-reboot
    .\repair.appx.ps1 -SetupPath "D:\setup.exe"

    # Allow automatic reboot at completion
    .\repair.appx.ps1 -SetupPath "D:\setup.exe" -NoReboot:$false

.NOTES
    Log   : C:\VDI_GPO_Logs\repair_appx.log
    Original one-liner (replaced): setup.exe /auto upgrade /Quiet /NoReboot /DynamicUpdate Disable /ShowOOBE none
    Run As: Local Administrator
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SetupPath,

    [bool]$NoReboot = $true,

    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$LogDir  = "C:\VDI_GPO_Logs"
$LogFile = "$LogDir\repair_appx_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $line = "[$(Get-Date -Format 'HH:mm:ss')] [$Level] $Msg"
    $col  = switch ($Level) { "ERROR"{"Red"} "WARN"{"Yellow"} "SUCCESS"{"Green"} default{"Gray"} }
    Write-Host $line -ForegroundColor $col
    Add-Content -Path $LogFile -Value $line
}

# ── Guard: Administrator ───────────────────────────────────────────────────────
$id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$pr = New-Object System.Security.Principal.WindowsPrincipal($id)
if (-not $pr.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "ERROR: This script must be run as Administrator." -ForegroundColor Red
    exit 1
}

# ── Setup log dir ─────────────────────────────────────────────────────────────
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }

Write-Log "=== Windows In-Place Repair / Upgrade ===" "INFO"
Write-Log "Setup path : $SetupPath"
Write-Log "NoReboot   : $NoReboot"
Write-Log "DryRun     : $DryRun"

# ── Guard: Validate setup.exe ─────────────────────────────────────────────────
if (-not (Test-Path $SetupPath)) {
    Write-Log "ERROR: setup.exe not found at: $SetupPath" "ERROR"
    Write-Log "Mount the Windows 11 ISO first (e.g., Mount-DiskImage) or provide the extracted media path." "ERROR"
    exit 1
}

$setupItem = Get-Item $SetupPath -ErrorAction Stop
if ($setupItem.Name -ine "setup.exe") {
    Write-Log "ERROR: Specified path does not point to setup.exe — got: $($setupItem.Name)" "ERROR"
    exit 1
}

Write-Log "setup.exe found: $($setupItem.FullName)  ($([math]::Round($setupItem.Length/1MB,1)) MB)" "SUCCESS"

# ── Guard: Disk space (Windows Setup requires ~8 GB free) ─────────────────────
$sysDrive = $env:SystemDrive
$disk     = Get-PSDrive -Name ($sysDrive.TrimEnd(':')) -ErrorAction SilentlyContinue
if ($disk) {
    $freeGB = [math]::Round($disk.Free / 1GB, 1)
    Write-Log "Free disk space on $sysDrive : $freeGB GB" "INFO"
    if ($freeGB -lt 8) {
        Write-Log "ERROR: Less than 8 GB free on $sysDrive. Windows Setup requires at least 8 GB." "ERROR"
        exit 1
    }
}

# ── Build setup.exe arguments ─────────────────────────────────────────────────
$setupArgs = @(
    "/auto", "upgrade",
    "/Quiet",
    "/DynamicUpdate", "Disable",
    "/ShowOOBE", "none"
)
if ($NoReboot) { $setupArgs += "/NoReboot" }

$cmdDisplay = "`"$SetupPath`" " + ($setupArgs -join " ")
Write-Log "Command: $cmdDisplay" "INFO"

if ($DryRun) {
    Write-Log "DRY RUN — command not executed. Review parameters above then remove -DryRun to proceed." "WARN"
    exit 0
}

# ── Confirmation prompt ────────────────────────────────────────────────────────
Write-Host ""
Write-Host "CAUTION: In-place Windows upgrade/repair will begin." -ForegroundColor Yellow
Write-Host "  This modifies system files and may reset some settings." -ForegroundColor Yellow
Write-Host "  Ensure you have a snapshot/checkpoint before proceeding." -ForegroundColor Yellow
Write-Host ""
$confirm = Read-Host "Type YES to continue, anything else to abort"
if ($confirm -ne "YES") {
    Write-Log "Aborted by user." "WARN"
    exit 0
}

# ── Execute ────────────────────────────────────────────────────────────────────
Write-Log "Starting Windows Setup (this may take 30-60 minutes)..." "INFO"
try {
    $proc = Start-Process -FilePath $SetupPath -ArgumentList $setupArgs `
        -Wait -PassThru -NoNewWindow

    Write-Log "Setup.exe exited with code: $($proc.ExitCode)" "INFO"

    switch ($proc.ExitCode) {
        0       { Write-Log "Setup completed successfully." "SUCCESS" }
        3010    { Write-Log "Setup completed — reboot required." "WARN" }
        default { Write-Log "Setup returned exit code $($proc.ExitCode). Check Windows Setup logs at C:\`$WINDOWS.~BT\Sources\Panther\setupact.log" "WARN" }
    }
} catch {
    Write-Log "Setup.exe failed to launch: $_" "ERROR"
    exit 1
}

Write-Log "Log: $LogFile"
Write-Log "Windows Setup logs: C:\`$WINDOWS.~BT\Sources\Panther\setupact.log"
if ($NoReboot) {
    Write-Log "REMINDER: Reboot manually before sealing or re-running Build scripts." "WARN"
}
