# MSP Audit Toolkit - Packaging Instructions

## Files in this package
- `MSPAuditToolkit.ps1` - the actual GUI application (WinForms, two tabs: Tenant Setup / Run Audit)
- `Launch.cmd` - hidden launcher that starts the GUI with no console window
- `Package.sed` - IExpress config, pre-filled to bundle the two files above

## One-time machine requirement
The target machine needs PowerShell 5.1+ (built into Windows 10/11) and the Microsoft.Graph
module. IExpress does not compile PowerShell into a native executable — it just bundles files
into a self-extracting archive and runs a launch command on extraction. If Microsoft.Graph isn't
installed, the GUI will pop a message telling the user to run:

    Install-Module Microsoft.Graph -Scope CurrentUser

## Building the EXE

### Option A - use the pre-filled SED file (fastest)
1. Put `MSPAuditToolkit.ps1`, `Launch.cmd`, and `Package.sed` in the same folder.
2. Open a command prompt in that folder and run:

       iexpress /N Package.sed

3. This produces `MSPAuditToolkit.exe` in the same folder — a single self-extracting
   file you can hand out or drop on a share.

### Option B - use the IExpress wizard (if you want to tweak anything)
1. Run `iexpress.exe` (built into Windows).
2. Choose **Create new Self Extraction Directive file** → Next.
3. Choose **Extract files and run an installation command** → Next.
4. Give it a package title (e.g. "MSP Audit Toolkit") → Next.
5. **No prompt** → Next.
6. **Do not display a license** → Next.
7. **Add files**: add both `MSPAuditToolkit.ps1` and `Launch.cmd` → Next.
8. **Install program**: choose `Launch.cmd` → Next.
9. **Show window**: Hidden → Next.
10. **Finished message**: No message → Next.
11. **Package name**: choose where to save `MSPAuditToolkit.exe`, check
    "Store files using long file name inside package" → Next.
12. **Configure restart**: No restart → Next.
13. Save the SED file if you want to rebuild later → Next → Finish.

## What the user experience looks like
Double-clicking `MSPAuditToolkit.exe` silently extracts the two files to a temp folder and
launches the GUI — no console window, no visible install steps. Note: because it extracts to
a temp folder, `Tenants.csv` / `TenantAuditApps.csv` / the audit report CSVs will be created
there too, not next to wherever the EXE itself is sitting. For a persistent setup, it's cleaner
to have techs extract it once to a permanent folder (e.g. `C:\MSPTools\`) and run
`MSPAuditToolkit.ps1` directly from there rather than re-running the EXE each time — that
way `Tenants.csv` and the credential/report CSVs persist between audit runs instead of living
in a temp folder that gets cleaned up.
