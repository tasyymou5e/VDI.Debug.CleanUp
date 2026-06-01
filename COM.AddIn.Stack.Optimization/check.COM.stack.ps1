# Requires -Version 7.0
# This script MUST be run as an Administrator.

# ==========================================
# 0. PRE-FLIGHT & CONFIGURATION
# ==========================================
Set-ExecutionPolicy Bypass -Scope Process -Force

Write-Host "--- OFFICE COM ADD-IN AUDIT & VDI OPTIMIZATION UTILITY ---" -ForegroundColor Cyan

# Mandatory Elevation Check
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Host "[X] CRITICAL ERROR: This script must be run with Administrator privileges!" -ForegroundColor Red
    if ($Host.Name -eq "ConsoleHost") { Read-Host "Press Enter to exit" }
    exit
}

# --- Logging Setup ---
$ScriptDir = $PSScriptRoot
$LogFile = Join-Path -Path $ScriptDir -ChildPath "COM_Add-in_Audit.log"
"--- COM ADD-IN AUDIT LOG: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ---" | Out-File -FilePath $LogFile -Force

function Write-Log {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
    $CleanMessage = $Message -replace "\[\✓\]|\[X\]|\[\!\]|\[i\]", "" | Out-String
    "$(Get-Date -Format 'HH:mm:ss') | $($CleanMessage.Trim())" | Out-File -FilePath $LogFile -Append
}

# ==========================================
# 1. DISCOVERY & AUDIT PHASE
# ==========================================
Write-Log -Message "`n--- PHASE 1: AUDITING ALL MACHINE-WIDE COM ADD-INS ---" -LogFile $LogFile -Color Yellow

$OfficeApps = @("Word", "Excel", "PowerPoint", "Outlook")
$BaseRegPath = "HKLM:\SOFTWARE\Microsoft\Office"
$FoundMisconfiguration = $false

foreach ($App in $OfficeApps) {
    $AddinsPath = "$BaseRegPath\$App\Addins"
    if (Test-Path $AddinsPath) {
        Write-Log -Message "`n[*] Checking Microsoft $App..." -LogFile $LogFile -Color Cyan
        $AddinKeys = Get-ChildItem -Path $AddinsPath
        
        if (-not $AddinKeys) {
            Write-Log -Message "  [i] No machine-wide add-ins found for this application." -LogFile $LogFile -Color Gray
            continue
        }

        foreach ($Key in $AddinKeys) {
            $AddinName = (Get-ItemProperty -Path $Key.PSPath -Name "FriendlyName" -ErrorAction SilentlyContinue).FriendlyName
            $LoadBehavior = (Get-ItemProperty -Path $Key.PSPath -Name "LoadBehavior" -ErrorAction SilentlyContinue).LoadBehavior
            
            if ($null -eq $AddinName) { $AddinName = $Key.PSChildName }

            if ($LoadBehavior -eq 3) {
                Write-Log -Message "  [✓] '$AddinName' is correctly set to Load at Startup (LoadBehavior=3)." -LogFile $LogFile -Color Green
            } else {
                Write-Log -Message "  [!] WARNING: '$AddinName' has a LoadBehavior of '$LoadBehavior'. It may not load correctly." -LogFile $LogFile -Color Yellow
                $FoundMisconfiguration = $true
            }
        }
    }
}

# ==========================================
# 2. VDI OPTIMIZATION (RESILIENCY) PHASE
# ==========================================
Write-Log -Message "`n--- PHASE 2: APPLYING VDI RESILIENCY SETTINGS FOR OUTLOOK ---" -LogFile $LogFile -Color Yellow

# ProgIDs of common enterprise add-ins prone to being disabled by Outlook.
$AddinsToProtect = @{
    "PDFMOutlook.PDFMOutlook"            = "Adobe Acrobat PDFMaker";
    "TeamsAddin.FastConnect"             = "Microsoft Teams Meeting Add-in";
    "Jabber.JabberOfficeIntegration.1"   = "Cisco Jabber / WebEx";
    "MSIP.OutlookAddin"                  = "Azure Information Protection";
    "VontuOfficeOL.Connect"              = "Broadcom/Symantec DLP";
    "dsig.Connect"                       = "DocuSign";
    "SalesforceForOutlook"               = "Salesforce"
}

$ResiliencyPath = "HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"

try {
    # Ensure the registry key path exists
    if (-not (Test-Path $ResiliencyPath)) {
        New-Item -Path $ResiliencyPath -Force -Recurse | Out-Null
        Write-Log -Message "  [i] Created 'DoNotDisableAddinList' registry key." -LogFile $LogFile -Color Gray
    }

    foreach ($ProgID in $AddinsToProtect.Keys) {
        $FriendlyName = $AddinsToProtect[$ProgID]
        # Set a DWORD value of 1 for each ProgID to prevent Outlook from disabling it.
        New-ItemProperty -Path $ResiliencyPath -Name $ProgID -Value 1 -PropertyType DWORD -Force -ErrorAction Stop | Out-Null
        Write-Log -Message "  [✓] Applied resiliency policy for '$FriendlyName'." -LogFile $LogFile -Color Green
    }
    Write-Log -Message "`n[SUCCESS] All targeted Outlook add-ins have been hardened against automatic disabling." -LogFile $LogFile -Color Cyan

} catch {
    Write-Log -Message "`n[X] FAILED to apply resiliency settings. Error: $($_.Exception.Message)" -LogFile $LogFile -Color Red
}

# ==========================================
# 3. CONCLUSION
# ==========================================
Write-Log -Message "`n--- AUDIT & OPTIMIZATION COMPLETE ---" -LogFile $LogFile -Color Cyan
if ($FoundMisconfiguration) {
    Write-Log -Message "[!] One or more add-ins have a non-standard LoadBehavior. Consider running a repair on the parent application or manually setting the LoadBehavior to 3." -LogFile $LogFile -Color Yellow
} else {
    Write-Log -Message "[✓] All discovered add-ins are configured to load at startup." -LogFile $LogFile -Color Green
}
