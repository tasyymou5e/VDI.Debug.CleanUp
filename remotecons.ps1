# Requires -Version 7.0
# Run as Administrator
#   C:\Temp\VMware-VMRC-<version>.exe /s /v"/qn ALLUSERS=1 EULAS_AGREED=1 AUTOSOFTWAREUPDATE=0 /norestart"

# 1. DEFINE CONFIGURATION
$InstallerPath = "C:\Temp\VMware-VMRC-<version>.exe" # Your specific filename
$VMRCPath = "C:\Program Files (x86)\VMware\VMware Remote Console\vmrc.exe"
$MinVCVersion = [version]"14.40.33816"

Write-Host "--- VMWARE REMOTE CONSOLE (VMRC) INSTALLATION (ALL USERS) ---" -ForegroundColor Cyan

# 2. PRE-INSTALL: CHECK FOR INSTALLER
if (-not (Test-Path $InstallerPath)) {
    Write-Host "[X] ERROR: Installer not found at $InstallerPath!" -ForegroundColor Red
    return
}

# 3. PRE-INSTALL: CHECK VISUAL C++ DEPENDENCY
Write-Host "`n[1/3] Checking Visual C++ Dependencies..." -ForegroundColor Yellow
$VCRedistInstalled = $false

# Registry paths for Visual C++ 2015-2022 Runtimes
$VCRegPaths = @(
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x86",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\x64"
)

foreach ($Path in $VCRegPaths) {
    if (Test-Path $Path) {
        $VersionStr = (Get-ItemProperty -Path $Path -Name "Version" -ErrorAction SilentlyContinue).Version
        if ($VersionStr) {
            $CleanVersion = $VersionStr.TrimStart('v') -replace '[a-zA-Z]',''
            try {
                $InstalledVersion = [version]$CleanVersion
                if ($InstalledVersion -ge $MinVCVersion) {
                    $VCRedistInstalled = $true
                    $Arch = if ($Path -match "x86") { "x86" } else { "x64" }
                    Write-Host "      [✓] Found VC++ ($Arch): $InstalledVersion" -ForegroundColor Green
                }
            } catch { }
        }
    }
}

if (-not $VCRedistInstalled) {
    Write-Host "      [X] ERROR: Missing Visual C++ 2015-2022 Redistributable!" -ForegroundColor Red
    Write-Host "      VMRC 13.0 requires version $MinVCVersion or higher." -ForegroundColor Gray
    return
}

# 4. EXECUTE MACHINE-WIDE INSTALLATION
Write-Host "`n[2/3] Starting Silent Installation for All Users..." -ForegroundColor Yellow

# Parameters Breakdown:
# /s = Silent mode for InstallShield
# /v = Pass arguments to internal MSI (Must be wrapped in double quotes)
# /qn = Quiet / No UI
# ALLUSERS=1 = Force machine-wide installation (Crucial for VDI)
# EULAS_AGREED=1 = Accept license
# AUTOSOFTWAREUPDATE=0 = Disable auto-updates
$InstallerArgs = '/s /v"/qn ALLUSERS=1 EULAS_AGREED=1 AUTOSOFTWAREUPDATE=0 /norestart"'

Write-Host "      Executing: $InstallerPath $InstallerArgs" -ForegroundColor Gray

# Start the InstallShield wrapper
$Process = Start-Process -FilePath $InstallerPath -ArgumentList $InstallerArgs -Wait -PassThru

Write-Host "      Waiting for background Windows Installer (msiexec) to finish..." -ForegroundColor Gray
# The EXE wrapper closes fast, so we loop to wait for the internal MSI process to write the files
while (Get-Process -Name "msiexec" -ErrorAction SilentlyContinue | Where-Object {$_.Path -match "syswow64"}) {
    Start-Sleep -Seconds 3
}

# 5. VERIFY RESULT DIRECTLY FROM FILE SYSTEM
Write-Host "`n[3/3] Verifying Application Path..." -ForegroundColor Yellow

if (Test-Path $VMRCPath) {
    $Version = (Get-Item $VMRCPath).VersionInfo.ProductVersion
    Write-Host "[✓] SUCCESS: VMRC Version $Version successfully installed in Program Files." -ForegroundColor Green
    
    # Optional: Verify it was registered in HKLM (All Users hive) and not HKCU
    $HKLMReg = Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object {$_.DisplayName -like "*VMware Remote Console*"}
    if ($HKLMReg) {
        Write-Host "      [✓] Confirmed: Registered in HKLM (All Users) registry hive." -ForegroundColor Green
    } else {
        Write-Host "      [!] WARNING: Files exist, but not found in HKLM Uninstall registry." -ForegroundColor Yellow
    }

} else {
    Write-Host "[X] FAILED: vmrc.exe not found at $VMRCPath." -ForegroundColor Red
    Write-Host "    * TROUBLESHOOTING: Check the log file at: %TEMP%\vminst.log for specific MSI failures." -ForegroundColor Gray
}

Write-Host "`n--- VMRC INSTALLATION COMPLETE ---" -ForegroundColor Cyan
