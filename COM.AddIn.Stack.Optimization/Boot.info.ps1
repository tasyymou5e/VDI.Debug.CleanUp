# Requires -Version 7.0
# Run as Administrator

# ==========================================
# 0. PRE-FLIGHT & CONFIGURATION
# ==========================================
Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

Write-Host "--- VDI BOOT PROCESS CAPTURE DEPLOYMENT ---" -ForegroundColor Cyan

$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Host "[X] ERROR: This script must be run as Administrator." -ForegroundColor Red
    exit
}

# Define paths
$TargetDir = "C:\Temp\VDI_Diagnostics"
$PayloadScript = Join-Path $TargetDir "CaptureBoot.ps1"
$LogFile = Join-Path $TargetDir "BootCapture.log"

if (-not (Test-Path $TargetDir)) { New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null }

# ==========================================
# 1. CREATE THE PAYLOAD SCRIPT
# ==========================================
Write-Host "`n[1/2] Creating boot capture payload script at $PayloadScript..." -ForegroundColor Yellow

$PayloadContent = @"
# --- VDI BOOT CAPTURE PAYLOAD ---
`$LogFile = `"$LogFile`"

`"==========================================`" | Out-File -FilePath `$LogFile -Append
`"BOOT CAPTURE INITIATED: `$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff'))`" | Out-File -FilePath `$LogFile -Append
`"USER CONTEXT: `$(whoami)`" | Out-File -FilePath `$LogFile -Append
`"==========================================`" | Out-File -FilePath `$LogFile -Append

`"--- RUNNING PROCESSES (Highest CPU First) ---`" | Out-File -FilePath `$LogFile -Append
Get-Process | Sort-Object CPU -Descending | Select-Object Name, Id, CPU, WorkingSet | Format-Table -AutoSize | Out-File -FilePath `$LogFile -Append

`"--- SERVICES CURRENTLY STARTING ---`" | Out-File -FilePath `$LogFile -Append
Get-Service | Where-Object { `$_.Status -eq 'StartPending' } | Select-Object Name, DisplayName | Format-Table -AutoSize | Out-File -FilePath `$LogFile -Append

`"--- ACTIVE SETUP KEYS EXECUTING ---`" | Out-File -FilePath `$LogFile -Append
Get-ItemProperty `"HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\*`" -ErrorAction SilentlyContinue | Select-Object PSChildName, StubPath | Format-Table -AutoSize | Out-File -FilePath `$LogFile -Append

`"CAPTURE COMPLETE: `$((Get-Date).ToString('HH:mm:ss.fff'))`n`" | Out-File -FilePath `$LogFile -Append
"@

$PayloadContent | Out-File -FilePath $PayloadScript -Force -Encoding UTF8
Write-Host "  [✓] Payload script created successfully." -ForegroundColor Green

# ==========================================
# 2. SCHEDULE THE PAYLOAD (Using schtasks.exe)
# ==========================================
Write-Host "`n[2/2] Registering payload to run at System Boot..." -ForegroundColor Yellow

$TaskName = "VDI_Boot_Diagnostic_Capture"

# Before creating, attempt to delete any existing task with the same name
schtasks.exe /Delete /TN $TaskName /F 2>$null | Out-Null

# Construct the schtasks.exe command
# /Create : Create a new task
# /TN     : Task Name
# /TR     : Task Run (The executable and arguments to run)
# /SC     : Schedule type (ONSTART means at boot)
# /RU     : Run As User ("SYSTEM" means run with highest privileges)
# /RL     : Run Level (HIGHEST)

$TaskRun = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PayloadScript`""
$SchTasksArgs = "/Create /TN `"$TaskName`" /TR `"$TaskRun`" /SC ONSTART /RU SYSTEM /RL HIGHEST /F"

# Execute schtasks.exe to create the task
$Process = Start-Process "schtasks.exe" -ArgumentList $SchTasksArgs -Wait -PassThru -WindowStyle Hidden

if ($Process.ExitCode -eq 0) {
    Write-Host "  [✓] Scheduled Task '$TaskName' registered successfully via schtasks." -ForegroundColor Green
} else {
    Write-Host "  [X] FAILED to register scheduled task. Exit code: $($Process.ExitCode)" -ForegroundColor Red
    exit
}

# ==========================================
# 3. CONCLUSION
# ==========================================
Write-Host "`n--- DEPLOYMENT COMPLETE ---" -ForegroundColor Cyan
Write-Host "1. Reboot the golden image." -ForegroundColor White
Write-Host "2. Wait for the system to fully boot to the logon screen." -ForegroundColor White
Write-Host "3. Log in as an administrator." -ForegroundColor White
Write-Host "4. Review the results in: $LogFile" -ForegroundColor Yellow
