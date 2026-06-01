# Requires -Version 7.0
# Run as Administrator

# ==========================================
# 0. PRE-FLIGHT & CONFIGURATION
# ==========================================
Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Host "[X] CRITICAL ERROR: This script must be run with Administrator privileges!" -ForegroundColor Red
    exit
}

# --- Logging Setup ---
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$LogFile   = Join-Path $ScriptDir "COMLoadFindings.log"

# Initialize fresh log file
"--- OFFICE COM ADD-IN VDI STATE AUDIT: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ---" | Out-File -FilePath $LogFile -Force

function Write-Log {
    param ([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
    # Clean ANSI codes and formatting symbols for the log file
    $CleanMsg = $Message -replace "\[\✓\]|\[X\]|\[\!\]|\[i\]|>>", "" | Out-String
    "$(Get-Date -Format 'HH:mm:ss') | $($CleanMsg.Trim())" | Out-File -FilePath $LogFile -Append
}

Write-Log "--- OFFICE COM ADD-IN VDI STATE AUDITOR ---" "Cyan"

# ==========================================
# 1. HELPER FUNCTIONS
# ==========================================
function Get-LoadBehaviorString([int]$BehaviorCode) {
    switch ($BehaviorCode) {
        0 { return "0 (Disconnected)" }
        1 { return "1 (Loaded - Unhandled)" }
        2 { return "2 (Load on Demand / Disabled)" }
        3 { return "3 (Load at Startup) [IDEAL]" }
        8 { return "8 (Load on Demand - User Selected)" }
        9 { return "9 (Load at Startup - User Selected)" }
        16 { return "16 (Connect First Time)" }
        default { return "$BehaviorCode (Unknown)" }
    }
}

# ==========================================
# 2. DEFINE SCOPES & APPLICATIONS
# ==========================================
$OfficeApps = @("Word", "Excel", "PowerPoint", "Outlook")
$Scopes = @{
    "Machine-Wide (HKLM)" = "HKLM:\SOFTWARE\Microsoft\Office";
    "Current User (HKCU)" = "HKCU:\SOFTWARE\Microsoft\Office"
}

# ==========================================
# 3. AUDIT EXECUTION
# ==========================================
$VDI_Critical_Addins = @("PDFMaker.OfficeAddin", "PDFMOutlook.PDFMOutlook", "TeamsAddin.FastConnect", "Jabber.JabberOfficeIntegration.1")
$Misconfigurations = 0

foreach ($ScopeName in $Scopes.Keys) {
    $BasePath = $Scopes[$ScopeName]
    Write-Log "`n==========================================" "Gray"
    Write-Log " SCANNING SCOPE: $ScopeName" "Yellow"
    Write-Log "==========================================" "Gray"

    foreach ($App in $OfficeApps) {
        $AddinsPath = "$BasePath\$App\Addins"
        
        if (Test-Path $AddinsPath) {
            $Addins = Get-ChildItem -Path $AddinsPath -ErrorAction SilentlyContinue
            
            if ($Addins.Count -gt 0) {
                Write-Log "`n>> Microsoft $App" "White"
                
                foreach ($Addin in $Addins) {
                    $ProgID = $Addin.PSChildName
                    $FriendlyName = (Get-ItemProperty -Path $Addin.PSPath -Name "FriendlyName" -ErrorAction SilentlyContinue).FriendlyName
                    $LoadBehavior = (Get-ItemProperty -Path $Addin.PSPath -Name "LoadBehavior" -ErrorAction SilentlyContinue).LoadBehavior
                    
                    if (-not $FriendlyName) { $FriendlyName = $ProgID }
                    $BehaviorStr = Get-LoadBehaviorString $LoadBehavior

                    # VDI Compliance Logic
                    if ($VDI_Critical_Addins -contains $ProgID) {
                        if ($LoadBehavior -eq 3) {
                            Write-Log "  [✓] $FriendlyName ($ProgID)" "Green"
                            Write-Log "      State: $BehaviorStr" "Gray"
                        } else {
                            Write-Log "  [X] $FriendlyName ($ProgID)" "Red"
                            Write-Log "      State: $BehaviorStr -> VDI CRITICAL ADD-IN SHOULD BE 3!" "Red"
                            $Misconfigurations++
                        }
                    } else {
                        # Non-critical add-ins
                        Write-Log "  [i] $FriendlyName ($ProgID)" "Cyan"
                        Write-Log "      State: $BehaviorStr" "Gray"
                    }
                }
            }
        }
    }
}

# ==========================================
# 4. SUMMARY
# ==========================================
Write-Log "`n--- AUDIT COMPLETE ---" "Cyan"
if ($Misconfigurations -gt 0) {
    Write-Log "[!] Found $Misconfigurations critical VDI add-in(s) in an incorrect load state." "Yellow"
    Write-Log "    Action: Run the Resiliency Hardening script to force LoadBehavior to 3." "Gray"
} else {
    Write-Log "[✓] All critical VDI add-ins are in the correct load state." "Green"
}
Write-Log "Detailed transcript saved to: $LogFile" "Gray"
