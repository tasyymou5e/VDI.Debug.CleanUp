# GPO.Reports — Audit & Reporting Tools

Generates auditable HTML reports of GPO settings and local security group configurations. Use during compliance review or to document the image configuration for stakeholders.

---

## Subfolders

### `GPO\` — GPO Settings Reporter

**`Get-GPOSettings.ps1`** — Interactive menu-driven auditor.

Reads Administrative Templates registry nodes for the following policy areas and produces per-node HTML reports:

**Computer policies:** FSLogix, Horizon Blast, DEM, Horizon Agent/Client, OneDrive, Start Menu/Taskbar

**User policies:** Horizon Blast, Edge, Teams, DEM, Horizon Agent/Client, OneDrive, Outlook, Start Menu/Taskbar

For each node it reads all registry values recursively and reports friendly formatted output (Enabled/Disabled/custom strings) alongside the raw registry path. Validates that each registry location exists before reporting.

`gpo-layout.txt` — Reference layout of all nodes and paths audited by the script.

---

### `Find.Security.Groups\` — Security Group & FSLogix Auditors

**`Find-SecurityGroupSources.ps1`** — Comprehensive audit of every source that creates or modifies local security groups on the image. Checks:
- Current local groups and their members
- Local GPO security template (`GptTmpl.inf` / Restricted Groups)
- LGPO, OSOT, Horizon provisioning scripts
- Startup scripts, scheduled tasks, and registry Run keys

Output: timestamped HTML report listing each finding by category, source, detail, and confidence level. Sample output: `SecurityGroupAudit_*.html`.

**`Get-FSLogixSettings.ps1`** — Dedicated FSLogix registry settings auditor. Reports container paths, VHD size, cloud cache configuration, volume type, and dynamic VHD settings.

**`Remediate-LGPOExe.ps1`** — Applies LGPO fixes identified during the security group audit.

`how.to.run.txt` / `how.to.fix.txt` — Usage and remediation instructions.

---

## When to Run

- During build validation to document the image configuration
- As part of compliance review cycles before image sealing
- When investigating unexpected local group membership or add-in policy conflicts
