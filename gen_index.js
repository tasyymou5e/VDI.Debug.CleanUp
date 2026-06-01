const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel,
  PageNumber, Header, Footer, VerticalAlign
} = require('docx');
const fs = require('fs');
const path = require('path');

const OUT = path.join(__dirname, 'Index.docx');

// ── colour palette ────────────────────────────────────────────────────────────
const DARK_BLUE  = '1A3A5C';
const MID_BLUE   = '1A6E9E';
const LIGHT_BLUE = 'D5E8F0';
const HEADER_BG  = '2C3E50';
const ROW_ALT    = 'EBF5FB';
const WHITE      = 'FFFFFF';

// ── helper: thin border ───────────────────────────────────────────────────────
const thinBorder = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
const allBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };

// ── helper: paragraph ─────────────────────────────────────────────────────────
const p = (text, opts = {}) =>
  new Paragraph({
    spacing: { before: opts.spaceBefore ?? 0, after: opts.spaceAfter ?? 80 },
    alignment: opts.align ?? AlignmentType.LEFT,
    children: [
      new TextRun({
        text,
        bold:      opts.bold      ?? false,
        italics:   opts.italic    ?? false,
        color:     opts.color     ?? '222222',
        size:      opts.size      ?? 20,
        font:      'Arial',
        break:     opts.break     ?? 0,
      })
    ]
  });

// ── helper: section heading ───────────────────────────────────────────────────
const sectionHeading = (text) =>
  new Paragraph({
    spacing: { before: 280, after: 120 },
    children: [
      new TextRun({ text, bold: true, color: WHITE, size: 24, font: 'Arial' })
    ],
    shading: { fill: DARK_BLUE, type: ShadingType.CLEAR },
    indent: { left: 120, right: 120 },
  });

// ── helper: script title bar ──────────────────────────────────────────────────
const scriptTitle = (filename, folder) =>
  new Paragraph({
    spacing: { before: 320, after: 60 },
    children: [
      new TextRun({ text: folder ? `${folder}  /  ` : '', color: MID_BLUE, size: 20, font: 'Arial' }),
      new TextRun({ text: filename, bold: true, color: DARK_BLUE, size: 24, font: 'Arial' }),
    ],
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE, space: 1 } },
  });

// ── helper: two-column info table ─────────────────────────────────────────────
const infoTable = (rows, shade) => {
  const tableRows = rows.map(([label, value], i) =>
    new TableRow({
      children: [
        new TableCell({
          borders: allBorders,
          width: { size: 2000, type: WidthType.DXA },
          shading: { fill: LIGHT_BLUE, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 18, font: 'Arial', color: DARK_BLUE })] })]
        }),
        new TableCell({
          borders: allBorders,
          width: { size: 7360, type: WidthType.DXA },
          shading: { fill: (shade && i % 2 === 1) ? ROW_ALT : WHITE, type: ShadingType.CLEAR },
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: value, size: 18, font: 'Arial' })] })]
        }),
      ]
    })
  );
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2000, 7360],
    rows: tableRows,
  });
};

// ── separator ─────────────────────────────────────────────────────────────────
const hr = () =>
  new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'CCCCCC', space: 1 } },
    children: [],
  });

// ═════════════════════════════════════════════════════════════════════════════
//  SCRIPT CATALOGUE
// ═════════════════════════════════════════════════════════════════════════════
const scripts = [

  // ── ROOT FOLDER ────────────────────────────────────────────────────────────
  {
    folder: null,
    file: 'altitude.and.appv.txt',
    type: 'PowerShell Script (PS7)',
    description:
      'Interactive menu-driven utility for installing the Omnissa App Volumes Agent and auditing ' +
      'Windows filter driver altitudes on a VDI golden image. Presents a numbered menu: ' +
      '(1) installs the App Volumes Agent MSI, configures the Horizon WaitAppVolumes registry key, ' +
      'and verifies the manager address and svservice registration; ' +
      '(2) runs a standalone altitude check of the svdriver, fslx, and vmm filter drivers using fltmc. ' +
      'All actions are logged to a timestamped .log file alongside the script.',
    prereqs: [
      'PowerShell 7.0 or later',
      'Run as Local Administrator',
      'App Volumes Agent MSI placed at C:\\Temp\\App_Volumes_Agent.msi',
      '$ManagerFQDN variable set to the App Volumes Manager FQDN (default placeholder: YOUR-APPVOL-MANAGER.domain.com)',
      'Omnissa Horizon Agent installed (for WaitAppVolumes registry path to exist)',
    ],
  },

  {
    folder: null,
    file: 'install.app.v.txt',
    type: 'PowerShell Script (PS7)',
    description:
      'Simplified (non-menu) App Volumes Agent installer. Runs a single silent MSI installation ' +
      'passing MANAGER_ADDR and MANAGER_PORT, verifies the manager address was written to the ' +
      'svservice registry key, confirms the svservice Windows service exists, and checks whether ' +
      'the svdriver filter driver is attached (expected absent until next reboot).',
    prereqs: [
      'PowerShell 7.0 or later',
      'Run as Local Administrator',
      'App Volumes Agent MSI at C:\\Temp\\App_Volumes_Agent.msi',
      '$ManagerFQDN set to your App Volumes Manager FQDN',
    ],
  },

  {
    folder: null,
    file: 'FSlogix.validate.txt',
    type: 'PowerShell Script (PS5.1+)',
    description:
      'Comprehensive Horizon golden-image provisioning and validation script. Configures four ' +
      'components required for silent Office activation and sign-in on instant clones: ' +
      '(1) Microsoft 365 Shared Computer Activation via registry; ' +
      '(2) Office ADAL / WAM identity settings; ' +
      '(3) FSLogix ODFC container validation (read-only unless -VHDLocations and/or -Force are passed); ' +
      '(4) Hybrid Azure AD Join (HAADJ) scheduled task with retry logic, fired at startup and logon. ' +
      'Fully idempotent — re-running on a configured image reports current state and changes nothing. ' +
      'Logs to C:\\ProgramData\\CompanyName\\Logs.',
    prereqs: [
      'PowerShell 5.1 or later',
      'Run as Local Administrator',
      'FSLogix Apps agent installed (version 2.9.8884.27471 or later recommended)',
      'Omnissa Horizon 8 environment (instant clones)',
      'On-premises AD synced to Entra ID via AAD Connect with True SSO',
      'Optional: -VHDLocations <UNC path> to write ODFC VHD location; -Force to overwrite existing value',
    ],
  },

  {
    folder: null,
    file: 'o365install.txt',
    type: 'PowerShell Script',
    description:
      'Offline Microsoft 365 Apps installation and VDI registry configuration script. Validates ' +
      'that setup.exe (ODT), M365.xml, and the pre-downloaded Office source folder are present, ' +
      'then runs a silent offline install using /configure. After installation, forcibly sets ' +
      'SharedComputerLicensing=1 and SCLCacheOverride=1 in the ClickToRun configuration registry ' +
      'key and verifies both values. All steps are logged to o365_install_log.txt.',
    prereqs: [
      'Run as Local Administrator',
      'Office Deployment Tool (setup.exe) in the script directory',
      'M365.xml configuration file in the script directory',
      'Pre-downloaded Office source folder (run officedownload.txt first)',
    ],
  },

  {
    folder: null,
    file: 'officedownload.txt',
    type: 'PowerShell Script',
    description:
      'Downloads the Microsoft 365 Apps source files via the Office Deployment Tool with a live ' +
      'progress indicator. Monitors the growing Office folder size every 2 seconds and displays a ' +
      'Write-Progress bar estimating percentage complete against a ~3300 MB target. Exits cleanly ' +
      'when the setup.exe download process finishes. Run this before o365install.txt.',
    prereqs: [
      'Run as Local Administrator',
      'Office Deployment Tool (setup.exe) in the script directory',
      'M365.xml configuration file defining the download scope',
      'Internet connectivity to Microsoft CDN (or internal WSUS/ODT mirror)',
    ],
  },

  {
    folder: null,
    file: 'teams.cleanup.txt',
    type: 'PowerShell Script (PS7)',
    description:
      'Teams (New) VDI provisioning script. Removes any existing MSTeams AppX packages using both ' +
      'a wrapped PowerShell 5.1 call (to avoid PS7 AppX serialization bugs) and the official ' +
      'teamsbootstrapper.exe -x cleanup. Provisions the new Teams MSIX for all users via ' +
      'teamsbootstrapper -p, installs the Outlook Meeting Add-in (TMA) via --installTMA, sets ' +
      'the IsWVDEnvironment=1 registry key, configures AppReadiness service to Manual start, ' +
      'and verifies the MSTeams package is staged correctly.',
    prereqs: [
      'PowerShell 7.0 or later',
      'Run as Local Administrator',
      'teamsbootstrapper.exe in C:\\Temp',
      'MSTeams-x64.msix in C:\\Temp',
    ],
  },

  {
    folder: null,
    file: 'validate.teams.and.optimization.txt',
    type: 'PowerShell Script',
    description:
      'Validates the Teams Outlook Meeting Add-in (TMA) installation across three checks: ' +
      '(1) Win32_Product MSI registration; ' +
      '(2) physical file presence of Microsoft.Teams.AddinLoader.dll under Program Files; ' +
      '(3) Outlook registry LoadBehavior value (must equal 3 for load at startup). ' +
      'Checks both 64-bit and 32-bit (WoW6432Node) registry paths and provides a fix command ' +
      'if LoadBehavior is incorrect.',
    prereqs: [
      'Run as Local Administrator (for Win32_Product query and HKLM registry read)',
      'Microsoft Teams already installed',
      'Microsoft 365 / Outlook installed',
    ],
  },

  {
    folder: null,
    file: 'remotecons.ps1',
    type: 'PowerShell Script (PS7)',
    description:
      'Installs VMware Remote Console (VMRC) silently for all users. Uses cmd.exe /c as a wrapper ' +
      'to pass the InstallShield /v switch without PowerShell quote-mangling. Waits for the ' +
      'background msiexec process (SysWOW64) to complete before verifying the vmrc.exe file ' +
      'exists and is registered in the HKLM Uninstall hive. Note: remotecons.txt is the ' +
      'same script saved as a plain-text reference copy.',
    prereqs: [
      'PowerShell 7.0 or later',
      'Run as Local Administrator',
      'VMware-VMRC-<version>.exe placed at C:\\Temp (rename variable $InstallerPath to match)',
      'Visual C++ 2015-2022 Redistributable (x86 and x64) installed',
    ],
  },

  {
    folder: null,
    file: 'additional.clean.txt',
    type: 'PowerShell Script',
    description:
      'Removes Microsoft Teams remnants from the Windows Default user profile so new VDI ' +
      'instant-clone users do not inherit stale Teams data. Deletes five known Teams folders under ' +
      'C:\\Users\\Default\\AppData. Also includes a registry block to load the Default NTUSER.DAT ' +
      'hive as HKU\\DefaultTemp, remove the legacy com.squirrel.Teams.Teams Run key, then unload ' +
      'the hive.',
    prereqs: [
      'Run as Local Administrator',
      'Executed on the golden image before sealing',
      'reg.exe (built-in) for hive load/unload',
    ],
  },

  {
    folder: null,
    file: 'per.user.cleanup.txt',
    type: 'PowerShell Script',
    description:
      'Golden-image cleanup script covering three areas: ' +
      '(1) removes a defined list of AppX/MSIX packages (Adobe Acrobat MSIX, Notepad++ MSIX, ' +
      'To Do, Power Automate, Bing Search, Cross Device, Dev Home) from all users and the ' +
      'provisioning manifest; ' +
      '(2) cleans leftover Teams and OneDrive directories from the Default profile; ' +
      '(3) disables VDI-optimization services (SysMain, WSearch, DiagTrack, dmwappushservice, ' +
      'MapsBroker, RetailDemo, SDRSVC, WbioSrvc). ' +
      'Also includes a one-liner to clear all Windows event logs.',
    prereqs: [
      'Run as Local Administrator',
      'PowerShell 5.1 or later with Appx module',
      'Executed on the golden image before sealing',
    ],
  },

  {
    folder: null,
    file: 'per.user.identify.txt',
    type: 'PowerShell Script',
    description:
      'Discovery script that scans the running golden image for software installed at the per-user ' +
      'level. Checks three sources: HKCU Uninstall registry keys (traditional MSI/EXE per-user ' +
      'installs), AppX/MSIX packages tied to specific user SIDs, and folders under ' +
      '%LOCALAPPDATA%\\Programs. Outputs a findings text file and an auto-generated remediation ' +
      'PowerShell script to C:\\Temp for review before execution.',
    prereqs: [
      'Run as Local Administrator (for -AllUsers AppX scan)',
      'PowerShell with Appx module (loads via -UseWindowsPowerShell fallback)',
    ],
  },

  {
    folder: null,
    file: 'one.drive.notepad.txt',
    type: 'Reference / Command Snippets',
    description:
      'Collection of winget and PowerShell commands for managing App Volumes and OneDrive on a ' +
      'golden image. Includes: uninstall commands for Notepad++ MSIX, OneDrive Sync MSIX, and the ' +
      'App Volumes Agent via winget; service/driver verification queries (Get-Service svservice, ' +
      'svdriver); winget and msiexec install commands for the App Volumes Agent; and a ' +
      'four-step connectivity verification block (TCP 443 test, service status, registry check, ' +
      'filter driver check) with $ManagerFQDN placeholder.',
    prereqs: [
      'winget (Windows Package Manager) available',
      'Run as Local Administrator',
      '$ManagerFQDN set to your App Volumes Manager FQDN before running verification block',
      'App Volumes Agent MSI available at C:\\Temp\\AppVolumesAgent.msi for install commands',
    ],
  },

  {
    folder: null,
    file: 'DisableFeatureUpdates.reg',
    type: 'Registry File (.reg)',
    description:
      'Windows Registry import file that pins Windows Update to the Windows 11 23H2 feature ' +
      'release, preventing automatic feature upgrades. Sets TargetReleaseVersion=1 and ' +
      'TargetReleaseVersionInfo="23H2" under the Windows Update policy key. Equivalent to ' +
      'the "Select the target feature update version" Group Policy setting.',
    prereqs: [
      'Run as Local Administrator',
      'Double-click the .reg file or run: reg import DisableFeatureUpdates.reg',
      'No scripting prerequisites — native registry import',
    ],
  },

  {
    folder: null,
    file: 'winget.txt',
    type: 'Diagnostic Output (Reference)',
    description:
      'Captured output of "winget list" from a Windows 11 24H2 golden image, sanitized of ' +
      'environment-specific telemetry (COAMS/DATT entries, hostname, serial number, and device ' +
      'identifiers have been redacted). Used as a reference inventory of installed applications ' +
      'and their winget IDs on a representative VDI image. Not a runnable script.',
    prereqs: [
      'Reference document only — no execution required',
      'To reproduce: run "winget list" as any user on a configured image',
    ],
  },

  {
    folder: null,
    file: 'appx.txt',
    type: 'Diagnostic Output (Reference)',
    description:
      'Captured output of "Get-AppxPackage | Select Name, InstallLocation" from a Windows 11 24H2 ' +
      'golden image, showing all installed AppX/MSIX packages and their installation paths. ' +
      'Used as a reference to identify packages for provisioning or removal decisions. ' +
      'Not a runnable script.',
    prereqs: [
      'Reference document only — no execution required',
      'To reproduce: run "Get-AppxPackage | Select Name, InstallLocation" as Administrator',
    ],
  },

  // ── COM.AddIn.Stack.Optimization FOLDER ───────────────────────────────────
  {
    folder: 'COM.AddIn.Stack.Optimization',
    file: 'check.COM.stack.ps1',
    type: 'PowerShell Script (PS7)',
    description:
      'Two-phase Office COM add-in audit and VDI resiliency tool. ' +
      'Phase 1 scans HKLM\\SOFTWARE\\Microsoft\\Office for all machine-wide add-ins across ' +
      'Word, Excel, PowerPoint, and Outlook, reporting each add-in\'s LoadBehavior value and ' +
      'flagging any that are not set to 3 (Load at Startup). ' +
      'Phase 2 applies Outlook add-in resiliency hardening by writing ' +
      'DWORD=1 entries under the DoNotDisableAddinList registry key for critical enterprise ' +
      'add-ins (Teams Meeting, Adobe PDFMaker, Cisco Jabber, AIP, Broadcom DLP, DocuSign, ' +
      'Salesforce), preventing Outlook from automatically disabling them on slow-load events. ' +
      'All findings logged to COM_Add-in_Audit.log in the script directory.',
    prereqs: [
      'PowerShell 7.0 or later',
      'Run as Local Administrator',
      'Microsoft Office 365 / 2016+ installed',
    ],
  },

  {
    folder: 'COM.AddIn.Stack.Optimization',
    file: 'COM.LoadBehavior.ps1',
    type: 'PowerShell Script (PS7)',
    description:
      'Detailed read-only COM add-in state auditor. Scans both HKLM (machine-wide) and ' +
      'HKCU (current-user) registry hives across Word, Excel, PowerPoint, and Outlook. ' +
      'For each add-in, decodes the LoadBehavior code into a human-readable string ' +
      '(0=Disconnected, 2=Load on Demand, 3=Load at Startup [IDEAL], etc.). ' +
      'Flags critical VDI add-ins (PDFMaker, TeamsAddin.FastConnect, Jabber) that are not ' +
      'at LoadBehavior=3. Outputs a full transcript to COMLoadFindings.log.',
    prereqs: [
      'PowerShell 7.0 or later',
      'Run as Local Administrator (for HKLM scope)',
      'Microsoft Office installed',
    ],
  },

  {
    folder: 'COM.AddIn.Stack.Optimization',
    file: 'Boot.info.ps1',
    type: 'PowerShell Script (PS7)',
    description:
      'Deploys a one-time boot diagnostic capture task to the golden image. Creates a payload ' +
      'script (CaptureBoot.ps1) at C:\\Temp\\VDI_Diagnostics that logs running processes sorted ' +
      'by CPU, services in StartPending state, and Active Setup stub keys — capturing the ' +
      'system state immediately after the next reboot. Registers the payload as a SYSTEM-level ' +
      'scheduled task (VDI_Boot_Diagnostic_Capture) using schtasks.exe /SC ONSTART. ' +
      'Review C:\\Temp\\VDI_Diagnostics\\BootCapture.log after rebooting.',
    prereqs: [
      'PowerShell 7.0 or later',
      'Run as Local Administrator',
      'Note: Script uses schtasks.exe (not PowerShell ScheduledTask cmdlets) for PS7 compatibility',
    ],
  },

  {
    folder: 'COM.AddIn.Stack.Optimization / HealthyAppX',
    file: 'appxrepait.txt',
    type: 'PowerShell Script (PS5.1)',
    description:
      'Emergency repair script for a broken Appx PowerShell module (Get-AppxPackage fails). ' +
      'Copies known-good Appx module files from a staging folder (C:\\ImagePrep\\HealthyAppx) ' +
      'over the broken system module directory, runs Unblock-File to remove the Mark of the Web ' +
      'from copied binaries, re-registers the three core Appx COM DLLs (AppxDeploymentClient.dll, ' +
      'AppxPackaging.dll, AppXDeploymentExtensions.OneCore.dll) via regsvr32, restarts AppXSvc, ' +
      'and reimports the Appx module. Use after OSOT-caused Appx corruption.',
    prereqs: [
      'PowerShell 5.1 (not PS7 — this repairs the PS5.1 module)',
      'Run as Local Administrator',
      'Healthy Appx module files staged at C:\\ImagePrep\\HealthyAppx (see ix web view.txt for extraction steps)',
    ],
  },

  {
    folder: 'COM.AddIn.Stack.Optimization / HealthyAppX',
    file: 'ix web view.txt',
    type: 'PowerShell Command Reference',
    description:
      'Step-by-step instructions for extracting healthy Appx PowerShell module files from a ' +
      'Windows 11 ISO/WIM. Uses Get-WindowsImage to identify the correct WIM index, ' +
      'Mount-WindowsImage to mount it read-only, copies the Appx module folder to ' +
      'C:\\ImagePrep\\HealthyAppx, then Dismount-WindowsImage -Discard. ' +
      'These extracted files are the source for appxrepait.txt.',
    prereqs: [
      'Run as Local Administrator',
      'Windows 11 ISO mounted or install.wim accessible (e.g., D:\\sources\\install.wim)',
      'C:\\Mount directory (created by the script)',
      'Sufficient disk space for WIM mounting (~6-8 GB)',
    ],
  },

  // ── GPO.Reports / Find.Security.Groups ────────────────────────────────────
  {
    folder: 'GPO.Reports / Find.Security.Groups',
    file: 'Find-SecurityGroupSources.ps1',
    type: 'PowerShell Script (PS7)',
    description:
      'Comprehensive audit script that identifies every possible source creating local security ' +
      'groups on a VDI golden image. Checks 12 categories: current local groups (Get-LocalGroup), ' +
      'Local GPO security template (GptTmpl.inf Restricted Groups), Local GPO Registry.pol ' +
      '(optionally parsed with lgpo.exe), lgpo.exe presence (broad search across C:\\), ' +
      'GPO startup scripts, common script locations, OSOT detection and template analysis, ' +
      'Horizon/VMware customization scripts and registry, non-Microsoft scheduled tasks, ' +
      'registry Run/RunOnce keys, Sysprep/SetupComplete/Unattend files, and Security event log ' +
      '(Event IDs 4731/4734/4728/4732 — group creation/deletion/membership changes). ' +
      'Outputs a color-coded HTML report with High/Medium/Low/Info confidence findings.',
    prereqs: [
      'PowerShell 7.0 or later',
      'Run as Local Administrator',
      'Security event log audit enabled (for Event ID section)',
      'Optional: lgpo.exe present on the machine for Registry.pol parsing',
      'Optional: -GroupNamesToSearch parameter to filter findings to specific group names',
    ],
  },

  {
    folder: 'GPO.Reports / Find.Security.Groups',
    file: 'Get-FSLogixSettings.ps1',
    type: 'PowerShell Script (PS7)',
    description:
      'Reads and documents all FSLogix registry settings from both the policy-managed path ' +
      '(HKLM:\\SOFTWARE\\Policies\\FSLogix) and the direct configuration path ' +
      '(HKLM:\\SOFTWARE\\FSLogix). Maps every known value to its human-readable name, expected ' +
      'values, and current status. Covers Profile Container, Office 365 Container (ODFC), ' +
      'Include/Exclude Groups, Logging, and Apps/General sections. Also reads the four FSLogix ' +
      'local security groups (Include/Exclude lists for Profiles and ODFC). ' +
      'Detects undocumented registry values not in the standard ADMX definition list. ' +
      'Outputs a color-coded HTML report saved alongside the script.',
    prereqs: [
      'PowerShell 7.0 or later',
      'Run as Local Administrator',
      'FSLogix Apps agent installed',
    ],
  },

  {
    folder: 'GPO.Reports / Find.Security.Groups',
    file: 'Remediate-LGPOExe.ps1',
    type: 'PowerShell Script (PS7)',
    description:
      'Parses the HTML report produced by Find-SecurityGroupSources.ps1, extracts every LGPO.exe ' +
      'file path found in High-confidence findings, backs each file up to a timestamped folder, ' +
      'and deletes it from the golden image. Includes a -WhatIfMode switch for a dry run that ' +
      'lists what would be deleted without making changes. Verifies all deletions succeeded. ' +
      'Produces a separate HTML remediation report showing status (Deleted, Backed Up, Skipped, ' +
      'Error) for each path. Run how.to.fix.txt for usage examples.',
    prereqs: [
      'PowerShell 7.0 or later',
      'Run as Local Administrator',
      'ExecutionPolicy set to RemoteSigned or Bypass',
      'SecurityGroupAudit HTML report from Find-SecurityGroupSources.ps1 (auto-detected from script folder)',
      'See how.to.fix.txt for dry-run and live-run command examples',
    ],
  },

  {
    folder: 'GPO.Reports / GPO',
    file: 'Get-GPOSettings.ps1',
    type: 'PowerShell Script (PS7)',
    description:
      'Menu-driven Administrative Templates settings auditor covering 15 GPO nodes across both ' +
      'Computer Configuration and User Configuration. Checks: FSLogix, Horizon Blast, Omnissa DEM, ' +
      'Omnissa Horizon Agent Configuration, Omnissa Horizon Client Configuration, OneDrive, and ' +
      'Start Menu and Taskbar (Computer); plus Horizon Blast, Microsoft Edge, Microsoft Teams, ' +
      'Omnissa DEM, Omnissa Horizon Agent Configuration, Omnissa Horizon Client Configuration, ' +
      'OneDrive, Outlook For Windows, and Start Menu and Taskbar (User). ' +
      'For each node, validates that registry locations exist (checking both legacy VMware and ' +
      'current Omnissa paths), reads all values recursively, detects policy conflicts where the ' +
      'same value exists under both a \\Policies\\ key and a direct key with different values, ' +
      'and generates a separate color-coded HTML report per node. ' +
      'Can be run interactively (menu), targeted (-Report <Key>), or in bulk (-Report All).',
    prereqs: [
      'PowerShell 7.0 or later',
      'Run as Local Administrator for full HKLM coverage',
      'Relevant ADMX templates applied (FSLogix, Omnissa Horizon, OneDrive)',
      'Use -List to print available report keys; -Report All to generate all reports at once',
    ],
  },

];

// ═════════════════════════════════════════════════════════════════════════════
//  BUILD DOCUMENT
// ═════════════════════════════════════════════════════════════════════════════

const children = [];

// ── Cover title ───────────────────────────────────────────────────────────────
children.push(
  new Paragraph({ spacing: { before: 0, after: 40 }, children: [] }),
  new Paragraph({
    spacing: { before: 0, after: 80 },
    alignment: AlignmentType.LEFT,
    children: [
      new TextRun({ text: 'VDI Golden Image', bold: true, size: 40, color: DARK_BLUE, font: 'Arial' }),
    ]
  }),
  new Paragraph({
    spacing: { before: 0, after: 60 },
    alignment: AlignmentType.LEFT,
    children: [
      new TextRun({ text: 'Script Index', bold: true, size: 56, color: DARK_BLUE, font: 'Arial' }),
    ],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: MID_BLUE, space: 1 } },
  }),
  p('Windows 11 24H2  |  Omnissa Horizon  |  FSLogix  |  Microsoft 365', {
    color: '555555', size: 18, spaceBefore: 60, spaceAfter: 20,
  }),
  p(`Generated: ${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}`, {
    color: '888888', size: 18, spaceAfter: 360,
  }),
);

// ── Purpose note ──────────────────────────────────────────────────────────────
children.push(
  new Paragraph({
    spacing: { before: 0, after: 100 },
    children: [
      new TextRun({ text: 'Purpose', bold: true, size: 22, color: DARK_BLUE, font: 'Arial' }),
    ],
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'CCCCCC', space: 1 } },
  }),
  p(
    'This index documents every script in the VDI golden-image build toolkit. ' +
    'Each entry describes what the script does, what it depends on, and what must be ' +
    'in place before running it. All scripts have been reviewed and sanitized of ' +
    'environment-specific identifiers.',
    { size: 19, spaceAfter: 280, color: '333333' }
  ),
);

// ── Section groupings ─────────────────────────────────────────────────────────
const sections = [
  { label: 'Root  —  Installation & Configuration', filter: s => !s.folder },
  { label: 'COM.AddIn.Stack.Optimization', filter: s => s.folder && s.folder.startsWith('COM.AddIn') },
  { label: 'GPO.Reports', filter: s => s.folder && s.folder.startsWith('GPO') },
];

for (const sec of sections) {
  children.push(sectionHeading(sec.label));

  for (const script of scripts.filter(sec.filter)) {
    children.push(
      scriptTitle(script.file, script.folder),
      infoTable([
        ['Type',        script.type],
        ['Description', script.description],
        ['Prerequisites',
          script.prereqs.map((r, i) => `${i + 1}.  ${r}`).join('\n')
        ],
      ], true),
      new Paragraph({ spacing: { before: 0, after: 120 }, children: [] }),
    );
  }
}

// ── Footer note ───────────────────────────────────────────────────────────────
children.push(
  hr(),
  p('All scripts require execution on the VDI golden image parent VM before sealing. ' +
    'Run as Local Administrator unless noted otherwise. ' +
    'Placeholders (YOUR-APPVOL-MANAGER.domain.com, C:\\Path\\To\\...) must be replaced with site-specific values.',
    { color: '666666', size: 17, italic: true }
  ),
);

// ═════════════════════════════════════════════════════════════════════════════
//  PACK & WRITE
// ═════════════════════════════════════════════════════════════════════════════

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: 'Arial', size: 20 } }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: 'VDI Golden Image  —  Script Index', color: '888888', size: 16, font: 'Arial' }),
              new TextRun({ children: ['\t'], font: 'Arial' }),
            ],
            tabStops: [{ type: 'right', position: 9360 }],
            border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: 'CCCCCC', space: 1 } },
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 2, color: 'CCCCCC', space: 1 } },
            children: [
              new TextRun({ text: 'Page ', color: '888888', size: 16, font: 'Arial' }),
              new TextRun({ children: [PageNumber.CURRENT], color: '888888', size: 16, font: 'Arial' }),
              new TextRun({ text: ' of ', color: '888888', size: 16, font: 'Arial' }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], color: '888888', size: 16, font: 'Arial' }),
            ]
          })
        ]
      })
    },
    children,
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log('Written:', OUT);
}).catch(err => {
  console.error(err);
  process.exit(1);
});
