# COM.AddIn.Stack.Optimization — Office Add-In & Appx Repair

Audits and repairs Office COM add-in registry configurations to prevent add-ins from being silently disabled during VDI sessions. Also repairs corrupted Appx packages.

---

## Scripts

### `check.COM.stack.ps1` — Main Auditor (Run This First)

Two-phase execution:

**Phase 1 — AUDIT**
Scans Word, Excel, PowerPoint, and Outlook for machine-wide COM add-ins. Checks that each add-in has `LoadBehavior=3` (auto-load). Flags any misconfigured or missing add-ins.

**Phase 2 — OPTIMIZE**
Applies VDI resiliency settings to prevent Outlook from disabling critical enterprise add-ins:
- Teams Meeting Add-in
- PDF Maker
- Jabber
- Azure Information Protection (AIP)

### `COM.LoadBehavior.ps1`

Standalone script to set `LoadBehavior` registry values for specific add-ins. Use when you need to target a single add-in without running the full audit.

### `repair.appx.ps1`

Repairs corrupted Appx packages using DISM and 7-Zip. Requires 7-Zip to be installed on the image. Use when event logs show Appx deployment errors.

### `Boot.info.ps1`

Boot-time diagnostics. Use to investigate slow startup or delayed logon issues.

---

## Supporting Files

| Path | Purpose |
|---|---|
| `HealthyAppX\` | Reference data (known-good Appx manifests) used by `repair.appx.ps1` — do not modify |
| `HealthyAppX\appxrepair.txt` | List of packages targeted for repair |

---

## When to Run

- Users report Office add-ins being disabled or grayed out after logon
- Event Viewer shows `LoadBehavior` being changed from 3 to 2 (disabled by Outlook)
- Appx deployment errors appear in the Application event log
- As a proactive step during golden image hardening, before the compliance audit
