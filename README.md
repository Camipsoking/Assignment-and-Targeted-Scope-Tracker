[README.md](https://github.com/user-attachments/files/33997446/README.md)
# Assignment and Targeted Scope Tracker

Assignment and Targeted Scope Tracker is a Windows desktop tool for auditing **Microsoft Intune assignments**. Pick an Entra ID group and Assignment and Targeted Scope Tracker lists 
every Intune object assigned to it — apps, configuration profiles, compliance policies, scripts, update rings, enrollment profiles, RBAC roles and more — including assignments the group 
inherits through nested (parent) groups.

The Intune admin center shows assignments one policy at a time. Assignment and Targeted Scope Tracker answers the reverse question: **"What does this group actually receive?"**

> Assignment and Targeted Scope Tracker runs entirely on the user's machine using delegated Microsoft Graph permissions. It is **read-only** and never changes anything in your tenant.

---

## Contents

- [Features](#features)
- [Use cases](#use-cases)
- [What gets audited](#what-gets-audited)
- [How it works](#how-it-works)
- [Repository layout](#repository-layout)
- [Requirements](#requirements)
- [Setup and installation](#setup-and-installation)
  - [1. Register the Entra ID application](#1-register-the-entra-id-application)
  - [2. Prepare the build workstation](#2-prepare-the-build-workstation)
  - [3. Customize the app](#3-customize-the-app)
  - [4. Build and sign](#4-build-and-sign)
  - [5. Bundle the Graph module](#5-bundle-the-graph-module)
  - [6. Test locally](#6-test-locally)
  - [7. Package for Intune](#7-package-for-intune)
  - [8. Deploy PowerShell 7](#8-deploy-powershell-7)
  - [9. Deploy AT-ST](#9-deploy-at-st)
- [Releasing updates](#releasing-updates)
- [Configuration reference](#configuration-reference)
- [Logs and troubleshooting](#logs-and-troubleshooting)
- [Security notes](#security-notes)
- [Disclaimer](#disclaimer)

---

## Features

- **Group-centric auditing** – select a group, see every Intune assignment that targets it.
- **Nested group resolution** – assignments made to parent groups are included and labeled `Inherited: <parent group>`.
- **Built-in targets** – audit `[All Users]` or `[All Devices]` directly from the group list.
- **Include / Exclude detection** – shows whether the group is included in or excluded from each assignment.
- **Result filters** – show or hide Direct and Inherited rows without re-running the audit.
- **Fast** – uses `$expand=assignments` where Graph supports it, and parallel lookups elsewhere.
- **Resilient** – retries throttled (429) and transient (503/504) Graph responses with backoff, and reports any gaps instead of silently dropping results.
- **Responsive UI** – audits run in the background; results stream into the grid as they're found.
- **Type-to-filter group search** with a cached group list and a Reload button.
- **Dark and light themes**, remembered between sessions.
- **No consent prompts for users** – sign-in uses only admin-consented permissions.
- **Signed and tamper-checked** – the launcher refuses to run an app script that isn't signed by the same certificate as the launcher.
- **Intune-ready packaging** – build script, install/uninstall scripts, and registry-based detection.

## Use cases

- **Troubleshooting** – "Why is this device getting this setting?" or "Why didn't this app install?"
- **Change validation** – confirm what a group will receive before adding users or devices to it.
- **Group cleanup** – check whether a group is still used by any Intune object before deleting it.
- **Security and compliance reviews** – document every policy that applies to a privileged or sensitive group.
- **Onboarding/offboarding** – verify a department or role group receives the expected apps and policies.
- **RBAC review** – see which Intune role assignments use a group as members or scope.
- **Tenant-wide baseline review** – audit `[All Users]` / `[All Devices]` to see what applies to everyone.

## What gets audited

| Area | Workloads |
|---|---|
| Apps | Applications (all platforms) |
| Configuration | Device configuration profiles, Settings Catalog (incl. newer Endpoint Security policies), Administrative Templates, legacy Endpoint Security templates |
| Compliance | Device compliance policies |
| Windows updates | Feature update, quality update and driver update profiles (update rings are covered by device configuration) |
| App protection / configuration | App protection (iOS, Android, Windows), app configuration for managed devices and managed apps |
| Scripts | Windows PowerShell scripts, Remediations, macOS shell scripts |
| Enrollment | Windows Autopilot deployment profiles, enrollment configurations/restrictions |
| Tenant administration | Scope tags, terms and conditions, policy sets, Intune RBAC role assignments |

Each result row shows **Workload**, **Name**, **Assignment** (Include/Exclude, or Member/Scope for RBAC), **Intent** (e.g. required/available), **Via** (Direct, Inherited, All Users, All Devices) and **Source** (Graph object type).

Adding another workload is a single line in the `$WorkloadManifest` array in the app script.

---

## How it works

```
AT-ST.exe  (compiled launcher, Windows PowerShell 5.1)
   │  verifies AT-ST.ps1 is signed by the same certificate
   │  finds pwsh.exe
   ▼
AT-ST.ps1  (the app, PowerShell 7, WPF UI)
   │  signs the user in (browser + MFA, cached afterward)
   │  uses the bundled Microsoft.Graph.Authentication module
   ▼
Microsoft Graph (delegated, read-only)
```

- **Why a launcher?** `ps2exe` compiles against Windows PowerShell 5.1, but AT-ST needs PowerShell 7 (parallel processing and reliable browser sign-in). The EXE is a small, signed launcher that starts the real app under `pwsh`.
- **Why delegated permissions?** App-only (service principal) access returns incomplete assignment data for several Intune workloads. Delegated access returns what the signed-in admin can see in the Intune portal, and every audit is attributable to a user in the Entra sign-in logs.
- **Sign-in** happens in the system browser (WAM is disabled because the app runs in a hidden PowerShell process). The token is cached, so later launches are silent until it expires.

---

## Repository layout

| File | Purpose |
|---|---|
| `AT-ST_IntuneAssignmentViewer_v5.6.ps1` | The application (PowerShell 7 + WPF). Becomes `AT-ST.ps1` in the build output. |
| `AT-ST_Launcher.ps1` | Launcher source. Compiled into `AT-ST.exe`. |
| `AT-ST-builder.ps1` | Build pipeline: generates the icon, compiles the launcher, signs both files, writes `version.txt`. |
| `New-ATSTIcon.ps1` | Optional: converts a PNG into a multi-size `AT-ST.ico`. |
| `config.json` | Template for tenant ID, client ID and throttle limit. |
| `Install-ATST.ps1` | Intune Win32 install script (runs as SYSTEM). |
| `Uninstall-ATST.ps1` | Intune Win32 uninstall script. |

Build output (`dist\`), which is what you package for Intune:

```
dist\
  AT-ST.exe            signed launcher
  AT-ST.ps1            signed app script
  AT-ST.ico            window / shortcut icon
  config.json          your tenant ID and client ID
  version.txt          app version (used for Intune detection)
  Install-ATST.ps1
  Uninstall-ATST.ps1
  Modules\Microsoft.Graph.Authentication\<version>\
```

---

## Requirements

**End-user devices**
- Windows 10 or 11, 64-bit
- PowerShell 7.2 or later (deployed as an Intune dependency — see step 8)
- An account assigned to the AT-ST enterprise application, with an Intune role that can read the data being audited

**Build workstation**
- Windows PowerShell 5.1 and PowerShell 7.2+
- `ps2exe` module (the builder installs it if missing)
- A code-signing certificate (an internal enterprise CA works well) in `Cert:\CurrentUser\My` or `Cert:\LocalMachine\My`
- [Microsoft Win32 Content Prep Tool](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool) (`IntuneWinAppUtil.exe`)

**Tenant**
- Permission to create an app registration and grant admin consent (Application Administrator or Global Administrator)

---

## Setup and installation

### 1. Register the Entra ID application

1. In the **Entra admin center**, go to **App registrations → New registration**.
   - Name: e.g. `Assignment and Targeted Scope Tracker`
   - Supported account types: **This organizational directory only**
2. Note the **Application (client) ID** and **Directory (tenant) ID** from the Overview page.
3. **Authentication**
   - **Add a platform → Mobile and desktop applications**, and add:
     - `http://localhost`
     - `https://login.microsoftonline.com/common/oauth2/nativeclient`
   - Under **Advanced settings**, set **Allow public client flows** to **Yes**. Save.
4. **API permissions → Add a permission → Microsoft Graph → Delegated permissions**, and add:

   | Permission | Used for |
   | `User.Read` | Sign-in |
   | `Group.Read.All` | Group list and nested group lookup |
   | `DeviceManagementApps.Read.All` | Apps, app protection, app configuration, policy sets |
   | `DeviceManagementConfiguration.Read.All` | Configuration, compliance, update profiles, Endpoint Security |
   | `DeviceManagementManagedDevices.Read.All` | Device-related policy data |
   | `DeviceManagementServiceConfig.Read.All` | Autopilot, enrollment configurations, terms and conditions |
   | `DeviceManagementRBAC.Read.All` | RBAC role assignments, scope tags |
   | `DeviceManagementScripts.Read.All` | PowerShell scripts, remediations, macOS shell scripts |
   | `offline_access`, `openid`, `profile` | Token refresh and sign-in |

5. Click **Grant admin consent for \<tenant\>** and confirm every permission shows **Granted**.
6. In **Enterprise applications → AT-ST → Properties**, set **Assignment required** to **Yes**, then add the users or groups who should use the tool under **Users and groups**.
7. *(Recommended)* Apply a Conditional Access policy to the app requiring MFA and a compliant device.

> Assignment and Targeted Scope Tracker requests only admin-consented permissions, so users never see a consent prompt. If a permission is missing, the affected workloads are skipped and listed in `audit.log`.

### 2. Prepare the build workstation

Put the repository files in one folder (avoid spaces in the path), then in a PowerShell window:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Get-ChildItem *.ps1 | Unblock-File
```

Find your code-signing certificate thumbprint:

```powershell
Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Format-List Subject, Thumbprint, NotAfter
```

### 3. Customize the app

1. **`config.json`** – fill in your values:

   ```json
   {
       "TenantId": "00000000-0000-0000-0000-000000000000",
       "EnterpriseAppClientId": "00000000-0000-0000-0000-000000000000",
       "ParallelThrottleLimit": 10
   }
   ```

2. **Logo (optional)** – the header logo is embedded as Base64 in the app script. Convert a PNG and paste the result between the `$LogoBase64 = @'` and `'@` markers, replacing `PASTE_YOUR_BASE64_HERE`:

   ```powershell
   [Convert]::ToBase64String([IO.File]::ReadAllBytes(".\logo.png")) | Set-Clipboard
   ```

3. **Icon (optional)** – to use a dedicated icon, save it as `AT-ST-icon.png` and run `.\New-ATSTIcon.ps1` to create `AT-ST.ico`. If you skip this, the builder generates an icon from the embedded logo.

4. **Branding** – set `-Company` when running the builder (step 4) and `$Publisher` in `Install-ATST.ps1` to your organization's name.

### 4. Build and sign

Run a full build the first time (and whenever the launcher or signing certificate changes):

```powershell
.\AT-ST-builder.ps1 `
    -CertThumbprint "YOUR-THUMBPRINT" `
    -AppScript ".\AT-ST_IntuneAssignmentViewer_v5.6.ps1" `
    -Company "Your Organization" `
    -IconFile ".\AT-ST.ico"
```

The builder:
- reads the version from `$AppVersion` in the app script and writes it to `dist\version.txt`
- compiles the launcher to `dist\AT-ST.exe` with your icon and version metadata
- copies the app script to `dist\AT-ST.ps1`
- signs both files (SHA256, timestamped) and verifies they share the same signer
- copies `config.json` into `dist\` (an existing `dist\config.json` is kept unless you pass `-ForceConfig`)

For an unsigned test build, add `-NoSign`. Unsigned launchers skip the signature check; **don't deploy unsigned builds**.

Then copy the remaining files into `dist\`:

```powershell
Copy-Item ".\AT-ST.ico", ".\Install-ATST.ps1", ".\Uninstall-ATST.ps1" ".\dist\"
```

### 5. Bundle the Graph module

Bundling `Microsoft.Graph.Authentication` with the app means users don't need it installed, and the app always uses the version you tested. Run in **PowerShell 7**:

```powershell
Save-Module Microsoft.Graph.Authentication -RequiredVersion "<tested version>" -Path ".\dist\Modules" -Repository PSGallery
```

Use the same version you tested with (`Get-Module -ListAvailable Microsoft.Graph.Authentication`).

### 6. Test locally

1. Run `dist\AT-ST.exe`.
2. Sign in when the browser opens, and complete MFA.
3. Pick a group and click **Run Audit**.
4. Check `C:\ProgramData\IntuneAssignmentViewer\audit.log`:
   - `Using Graph module: ...\dist\Modules\...` confirms the bundled module loaded
   - lines ending `via $expand` or `via per-item lookups` show how each workload was read
   - there should be no `FAILED` lines
5. Close and relaunch — this time there should be no sign-in prompt.

To force a fresh sign-in for testing, run in PowerShell 7:

```powershell
Connect-MgGraph -TenantId "<tenant-id>" -ClientId "<client-id>" -NoWelcome
Disconnect-MgGraph
```

### 7. Package for Intune (optional)

```powershell
.\IntuneWinAppUtil.exe -c ".\dist" -s "Install-ATST.ps1" -o ".\package"
```

### 8. Deploy PowerShell 7

Skip this if PowerShell 7 is already deployed to the target devices.

1. Download the x64 MSI from the [PowerShell releases page](https://github.com/PowerShell/PowerShell/releases).
2. Wrap it with `IntuneWinAppUtil.exe` and create a **Windows app (Win32)**:

   | Setting | Value |
   |---|---|
   | Install command | `msiexec /i "PowerShell-7.x.x-win-x64.msi" /qn ADD_PATH=1` |
   | Uninstall command | `msiexec /x {product-code} /qn` |
   | Install behavior | System |
   | Detection rule | File exists: `C:\Program Files\PowerShell\7\pwsh.exe` |

### 9. Deploy AT-ST

In Intune, create a **Windows app (Win32)** and upload `package\Install-ATST.intunewin`:

| Setting | Value |
|---|---|
| Install command | `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-ATST.ps1` |
| Uninstall command | `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Uninstall-ATST.ps1` |
| Install behavior | System |
| Device restart behavior | No specific action |
| Requirements | Windows 10/11, 64-bit |
| Detection rule | Registry · Key `HKEY_LOCAL_MACHINE\SOFTWARE\AT-ST` · Value `Version` · **Version comparison** · **Greater than or equal to** · `X.Y.Z` · Associated with a 32-bit app on 64-bit clients: **No** |
| Dependencies | PowerShell 7 app, automatically install: **Yes** |
| Assignment | Required or Available, to the same users/groups assigned to the enterprise app |

The install script:
- copies the app to `C:\Program Files\AT-ST` (replacing any previous version)
- creates `C:\ProgramData\IntuneAssignmentViewer` and gives standard users Modify access (logs, group cache, theme)
- adds a Start menu shortcut for all users
- writes `HKLM\SOFTWARE\AT-ST\Version` from `version.txt` for detection
- registers AT-ST in **Apps & features**

**Before rolling out widely:** pilot with one or two users, and give your security team the public code-signing certificate (`.cer`) if your Defender, WDAC or AppLocker policies need a publisher-based allow rule.

---

## Releasing updates

1. Edit the app script and bump `$AppVersion` (e.g. `5.6.0` → `5.7.0`).
2. Rebuild the script only — the EXE doesn't change:
   ```powershell
   .\AT-ST-builder.ps1 -CertThumbprint "YOUR-THUMBPRINT" -AppScript ".\AT-ST_IntuneAssignmentViewer_v5.7.ps1" -ScriptOnly
   ```
3. Re-run `IntuneWinAppUtil.exe`.
4. In Intune, upload the new `.intunewin` to the existing app and update the detection rule's version.

Devices on the older version no longer match the detection rule and update automatically.

Do a **full build** (without `-ScriptOnly`) if you change the launcher or renew the signing certificate. The launcher only runs scripts signed by its own certificate, so both files must be re-signed with the new certificate.

If an update needs a **new Graph permission**, add it to the app registration and grant admin consent *before* deploying.

---

## Configuration reference

**`config.json`** (searched next to the app first, then in `C:\ProgramData\IntuneAssignmentViewer\`)

| Key | Required | Description |
|---|---|---|
| `TenantId` | Yes | Directory (tenant) ID (GUID) |
| `EnterpriseAppClientId` | Yes | Application (client) ID of the app registration (GUID) |
| `ParallelThrottleLimit` | No | Concurrent per-item Graph lookups, 1–30. Default 10. Values of 10–15 work well; higher values risk throttling. Changes take effect at next launch, no rebuild needed. |

If no config file is found, AT-ST creates a template in ProgramData and tells the user where it is.

**Files written at runtime** – `C:\ProgramData\IntuneAssignmentViewer\`

| File | Contents |
|---|---|
| `audit.log` | Sign-in, workload methods, failures |
| `launcher.log` | Launcher signature checks and startup |
| `groups_cache.json` | Cached group list (refresh with **Reload Groups**) |
| `theme.json` | Theme preference |

---

## Logs and troubleshooting

| Log | Location |
|---|---|
| App | `C:\ProgramData\IntuneAssignmentViewer\audit.log` |
| Launcher | `C:\ProgramData\IntuneAssignmentViewer\launcher.log` |
| Install / uninstall | `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\AT-ST-Install.log` |
| Intune agent | `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\AppWorkload.log` |

| Symptom | Likely cause / fix |
|---|---|
| Consent prompt listing many permissions | One permission isn't admin-consented, or `config.json` has the wrong client ID. Grant consent and check the **Permissions** tab on the enterprise app. |
| `AADSTS500113: No reply address is registered` | Redirect URIs missing — see step 1.3. |
| `A window handle must be configured` | You're running an older build or from PowerShell ISE. Use a current build and run from `AT-ST.exe` or a normal PowerShell 7 window. |
| `ClientId is not a GUID` | Wrong value in `config.json` — use the **Application (client) ID** from App registrations, not the enterprise app's Object ID. |
| "Application Signature Invalid" on launch | `AT-ST.ps1` was edited or signed with a different certificate than the EXE. Rebuild with the builder. |
| "PowerShell 7 Required" | PowerShell 7 isn't installed — check the Intune dependency. |
| Status shows "Audit complete with gaps" | Some workloads or items couldn't be read. `FAILED` lines in `audit.log` give the reason (usually a missing permission or a feature your tenant doesn't use). |
| Install fails: `missing required file: version.txt` | Build with `AT-ST-builder.ps1`, which creates it, or add a file containing just the version number. |
| Window/taskbar shows the PowerShell icon | `AT-ST.ico` isn't next to `AT-ST.ps1`. |
| A group I expect isn't in the list | Click **Reload Groups** to refresh the cached group list. |

---

## Security notes

- **Read-only.** Assignment and Targeted Scope Tracker only issues `GET` requests to Microsoft Graph.
- **No secrets stored.** Users sign in interactively; there's no client secret or certificate credential in the app or config.
- **Delegated access.** Users see only what their own Intune role allows, and every sign-in appears under the enterprise app's sign-in logs.
- **Signature enforcement.** The signed launcher verifies the app script's signature and signer before running it.
- **Token cache.** Tokens are cached per user by the Microsoft Graph PowerShell SDK, as with any interactive Graph PowerShell session.

---

## Disclaimer

Assignment and Targeted Scope Tracker (AT-ST) is a community project and is not affiliated with or supported by Microsoft. Test in your environment before broad deployment.

Contributions, issues and suggestions are welcome.
