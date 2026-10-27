#Requires -Version 7.2
###############################################################################
# Assignment & Targeted Scope Tracker
# Intune Assignment Viewer - Imperial Edition (v5.6 / production)
#
# v5.6 changes (deployment):
#   - Uses the Microsoft.Graph.Authentication module bundled in the app's
#     Modules folder when present (version-locked; no per-user install, and
#     unaffected by other module versions on the machine). Falls back to an
#     installed module if no bundled copy exists.
#
# v5.5 changes (performance + reliability):
#   - Each workload first requests ?$expand=assignments, returning all items
#     with their assignments in a few paged calls. Workloads that reject
#     $expand fall back to per-item parallel lookups automatically.
#   - Retry with backoff on 429/503/504 for every Graph call (Retry-After
#     honored; otherwise exponential backoff with jitter, up to 5 attempts)
#   - Items/workloads that still fail are logged, and the status bar reports
#     "Audit complete with gaps" instead of silently dropping them
#   - audit.log records which method (expand / per-item) each workload used
#
# v5.4 changes:
#   - All Users and All Devices are now selectable entries at the top of the
#     group dropdown. Auditing them lists what's assigned to All Users / All
#     Devices (including RBAC scope).
#   - Regular group audits no longer list All Users / All Devices rows, so
#     results only show what targets the group (Direct + Inherited).
#   - Filter bar reduced to Direct / Inherited.
#
# v5.3 changes:
#   - Result filters: show/hide Direct, Inherited, All Users, All Devices
#     rows (live, no re-run needed) with a "Showing X of Y" count
#   - Fix: Run Audit / Reload stayed disabled after the first audit
#   - Window/taskbar icon set from AT-ST.ico (falls back to header logo)
#   - Sign-in requests .default (admin-consented permissions only), so users
#     never see a consent prompt; missing permissions are logged instead
#
# v5.2 changes:
#   - Effective-assignment auditing: direct group, inherited via parent
#     (nested) groups, and All Users / All Devices assignments (VIA column)
#   - 15 additional workloads (Windows Update profiles, App Protection,
#     App Configuration, legacy Endpoint Security, Scripts/Remediations,
#     Scope Tags, Terms and Conditions, Policy Sets)
#   - Delegated scope added: DeviceManagementScripts.Read.All
#   - Fix: theme toggle no longer clips "Switch to Imperial"
#   - Fix: selecting a group no longer leaves text highlighted
#
# v5 changes (packaging & maintainability):
#   - Externalized configuration (config.json next to app; no secrets in code)
#   - Visible app version (synced to EXE FileVersion + Intune app version)
#   - Workload manifest structure (adding a workload = one array entry)
#   - Runs on PowerShell 7 (superior browser auth; ForEach-Object -Parallel)
#   - First-run config template creation with clear guidance
#
# Retains:
#   - Star Wars themed WPF UI (Imperial dark / Rebel light) + persistent toggle
#   - Embedded AT-ST logo (self-contained)
#   - Delegated auth via Enterprise App (browser interactive; device-code fallback)
#   - Cached token reuse (near-silent after first sign-in)
#   - Background runspace orchestration (UI never freezes)
#   - Parallel Graph calls (throttled)
#   - Incremental result streaming
#   - Fast CollectionView-based, virtualized dropdown filtering
#   - Full coverage: Apps, Config, Compliance, Enrollment, RBAC
#
# Requires: PowerShell 7.2+, Microsoft.Graph.Authentication 2.19.0+
###############################################################################

###############################################################################
# APP CONSTANTS
###############################################################################

$AppVersion   = "5.6.0"
$AppName      = "AT-ST"
$AppFullName  = "Assignment and Targeted Scope Tracker"

###############################################################################
# RESOLVE APP ROOT (works as .ps1 and as compiled .exe)
###############################################################################

if ($PSScriptRoot) {
    $AppRoot = $PSScriptRoot
} elseif ($PSCommandPath) {
    $AppRoot = Split-Path -Parent $PSCommandPath
} else {
    try {
        $AppRoot = Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
    } catch {
        $AppRoot = (Get-Location).Path
    }
}

###############################################################################
# ASSEMBLIES + MODULE SAFETY CHECK
###############################################################################

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Prefer the Graph module bundled with the app (Modules\Microsoft.Graph.Authentication\<version>).
# Loading by full path guarantees every runspace uses the same tested version.
$GraphModuleRef = "Microsoft.Graph.Authentication"
$bundledGraphRoot = Join-Path $AppRoot "Modules\Microsoft.Graph.Authentication"
if (Test-Path $bundledGraphRoot) {
    $bundledPsd1 = Get-ChildItem -Path $bundledGraphRoot -Filter "Microsoft.Graph.Authentication.psd1" -Recurse -ErrorAction SilentlyContinue |
        Sort-Object { $v = $null; if ([version]::TryParse($_.Directory.Name, [ref]$v)) { $v } else { [version]"0.0" } } -Descending |
        Select-Object -First 1
    if ($bundledPsd1) { $GraphModuleRef = $bundledPsd1.FullName }
}

if ($GraphModuleRef -eq "Microsoft.Graph.Authentication" -and -not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    [System.Windows.MessageBox]::Show(
        "The Microsoft.Graph.Authentication module is not installed.`n`nRun this in PowerShell:`n`nInstall-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force",
        "$AppName - Missing Module",
        "OK",
        "Error"
    ) | Out-Null
    return
}

Import-Module $GraphModuleRef -ErrorAction Stop

###############################################################################
# PATHS (per-user state) + CONFIG (externalized)
###############################################################################

$AppDataDir = "$env:ProgramData\IntuneAssignmentViewer"
$CachePath  = "$AppDataDir\groups_cache.json"
$LogPath    = "$AppDataDir\audit.log"
$ThemePath  = "$AppDataDir\theme.json"

# Config is searched next to the app first (admin-editable, deploy-friendly),
# then falls back to ProgramData.
$ConfigCandidates = @(
    (Join-Path $AppRoot "config.json"),
    (Join-Path $AppDataDir "config.json")
)

function Write-Log {
    param ([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    if (-not (Test-Path $AppDataDir)) {
        New-Item -ItemType Directory -Path $AppDataDir -Force | Out-Null
    }
    Add-Content -Path $LogPath -Value "$timestamp | $Message"
}

function New-ConfigTemplate {
    param ([string]$Path)
    $template = [ordered]@{
        TenantId              = "REPLACE-WITH-TENANT-GUID"
        EnterpriseAppClientId = "REPLACE-WITH-APP-CLIENT-GUID"
        ParallelThrottleLimit = 10
    }
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $template | ConvertTo-Json -Depth 4 | Set-Content -Path $Path -Encoding UTF8
    Write-Log "Created config template at $Path"
}

function Load-Config {
    $found = $null
    foreach ($c in $ConfigCandidates) {
        if (Test-Path $c) { $found = $c; break }
    }

    if (-not $found) {
        # Create a template in ProgramData and guide the user
        $templatePath = Join-Path $AppDataDir "config.json"
        New-ConfigTemplate -Path $templatePath
        [System.Windows.MessageBox]::Show(
            "No configuration file was found.`n`nA template has been created at:`n$templatePath`n`nOpen it, fill in your TenantId and EnterpriseAppClientId, save, then relaunch $AppName.",
            "$AppName - Configuration Required",
            "OK",
            "Information"
        ) | Out-Null
        return $null
    }

    try {
        $cfg = Get-Content -Path $found -Raw | ConvertFrom-Json
    } catch {
        [System.Windows.MessageBox]::Show(
            "The configuration file could not be read:`n$found`n`n$($_.Exception.Message)",
            "$AppName - Configuration Error",
            "OK",
            "Error"
        ) | Out-Null
        return $null
    }

    # Validate required fields
    $missing = @()
    if ([string]::IsNullOrWhiteSpace($cfg.TenantId) -or $cfg.TenantId -like "REPLACE-*") { $missing += "TenantId" }
    if ([string]::IsNullOrWhiteSpace($cfg.EnterpriseAppClientId) -or $cfg.EnterpriseAppClientId -like "REPLACE-*") { $missing += "EnterpriseAppClientId" }

    if ($missing.Count -gt 0) {
        [System.Windows.MessageBox]::Show(
            "The configuration file is missing required values:`n$($missing -join ', ')`n`nEdit:`n$found`n`nThen relaunch $AppName.",
            "$AppName - Configuration Incomplete",
            "OK",
            "Warning"
        ) | Out-Null
        return $null
    }

    # Validate GUIDs
    $guidCheck = [ref]([guid]::Empty)
    if (-not [guid]::TryParse([string]$cfg.TenantId, $guidCheck)) {
        [System.Windows.MessageBox]::Show("TenantId in config is not a valid GUID.", "$AppName - Configuration Error", "OK", "Error") | Out-Null
        return $null
    }
    if (-not [guid]::TryParse([string]$cfg.EnterpriseAppClientId, $guidCheck)) {
        [System.Windows.MessageBox]::Show("EnterpriseAppClientId in config is not a valid GUID.", "$AppName - Configuration Error", "OK", "Error") | Out-Null
        return $null
    }

    $throttle = 10
    if ($cfg.ParallelThrottleLimit) {
        [int]::TryParse([string]$cfg.ParallelThrottleLimit, [ref]$throttle) | Out-Null
        if ($throttle -lt 1) { $throttle = 1 }
        if ($throttle -gt 30) { $throttle = 30 }
    }

    Write-Log "Loaded config from $found"

    return [PSCustomObject]@{
        TenantId              = [string]$cfg.TenantId
        EnterpriseAppClientId = [string]$cfg.EnterpriseAppClientId
        ParallelThrottleLimit = $throttle
        SourcePath            = $found
    }
}

$Config = Load-Config
if (-not $Config) { return }

Write-Log "Using Graph module: $GraphModuleRef"

$TenantId              = $Config.TenantId
$EnterpriseAppClientId = $Config.EnterpriseAppClientId
$ParallelThrottleLimit = $Config.ParallelThrottleLimit

$RequiredScopes = @(
    "User.Read",
    "Group.Read.All",
    "DeviceManagementApps.Read.All",
    "DeviceManagementConfiguration.Read.All",
    "DeviceManagementManagedDevices.Read.All",
    "DeviceManagementServiceConfig.Read.All",
    "DeviceManagementRBAC.Read.All",
    "DeviceManagementScripts.Read.All"
)

###############################################################################
# WORKLOAD MANIFEST  (extensibility: add a workload = add an entry here)
#
# Type 'assignments' : list endpoint, then per-item /assignments lookup
# Type 'rbac'        : list endpoint, membership/scope check (no /assignments)
#
# IntentMode:
#   'appIntent'    -> use assignment.intent (apps)
#   'policyIntent' -> use assignment.intent (config)
#   'enforced'     -> literal "Enforced" (compliance)
#   'none'         -> blank (enrollment, rbac)
###############################################################################

$WorkloadManifest = @(
    @{ Name = "Application";              Type = "assignments"; IntentMode = "appIntent";    ListUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps" }
    @{ Name = "Device Configuration";     Type = "assignments"; IntentMode = "policyIntent"; ListUri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations" }
    @{ Name = "Settings Catalog";         Type = "assignments"; IntentMode = "policyIntent"; ListUri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" }
    @{ Name = "Administrative Templates"; Type = "assignments"; IntentMode = "policyIntent"; ListUri = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations" }
    @{ Name = "Compliance";               Type = "assignments"; IntentMode = "enforced";     ListUri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies" }
    @{ Name = "Autopilot Profile";        Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles" }
    @{ Name = "Enrollment Configuration"; Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations" }
    # --- Windows Updates ---
    @{ Name = "Feature Update Profile";   Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceManagement/windowsFeatureUpdateProfiles" }
    @{ Name = "Quality Update Profile";   Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceManagement/windowsQualityUpdateProfiles" }
    @{ Name = "Driver Update Profile";    Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceManagement/windowsDriverUpdateProfiles" }
    # --- App Protection ---
    @{ Name = "App Protection (iOS)";     Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceAppManagement/iosManagedAppProtections" }
    @{ Name = "App Protection (Android)"; Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections" }
    @{ Name = "App Protection (Windows)"; Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceAppManagement/windowsManagedAppProtections" }
    # --- App Configuration ---
    @{ Name = "App Config (Devices)";     Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations" }
    @{ Name = "App Config (Apps)";        Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceAppManagement/targetedManagedAppConfigurations" }
    # --- Endpoint Security (legacy templates) ---
    @{ Name = "Endpoint Security (Legacy)"; Type = "assignments"; IntentMode = "none";       ListUri = "https://graph.microsoft.com/beta/deviceManagement/intents" }
    # --- Scripts (requires DeviceManagementScripts.Read.All) ---
    @{ Name = "PowerShell Script";        Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts" }
    @{ Name = "Remediation";              Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts" }
    @{ Name = "macOS Shell Script";       Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceManagement/deviceShellScripts" }
    # --- Tenant Administration ---
    @{ Name = "Scope Tag";                Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceManagement/roleScopeTags" }
    @{ Name = "Terms and Conditions";     Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/v1.0/deviceManagement/termsAndConditions" }
    @{ Name = "Policy Set";               Type = "assignments"; IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceAppManagement/policySets" }
    @{ Name = "RBAC Assignment";          Type = "rbac";        IntentMode = "none";         ListUri = "https://graph.microsoft.com/beta/deviceManagement/roleAssignments" }
)

# Disable WAM so browser sign-in works without a parent window handle
# (required because the EXE launches PowerShell 7 hidden)
try { Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction Stop }
catch { Write-Log "Set-MgGraphOption not available: $($_.Exception.Message)" }
$env:AZURE_IDENTITY_DISABLE_INTERACTIVEBROWSERBROKER = "true"

###############################################################################
# THEME DEFINITIONS
###############################################################################

$Themes = @{
    Imperial = @{
        Background       = "#0A0A0C"
        Surface          = "#161619"
        Subtle           = "#1F1F24"
        Border           = "#2E2E35"
        TextPrimary      = "#E8E8EB"
        TextSecondary    = "#8A8A93"
        Accent           = "#B01D22"
        AccentHover      = "#8B1519"
        AccentPressed    = "#6B0F13"
        DataGridAltRow   = "#131317"
        ButtonHover      = "#1F1F24"
        ButtonPressed    = "#2A2A30"
        HeaderBackground = "#111114"
    }
    Rebel = @{
        Background       = "#F4EFE6"
        Surface          = "#FFFFFF"
        Subtle           = "#EAE3D4"
        Border           = "#C8BFAE"
        TextPrimary      = "#1F2126"
        TextSecondary    = "#6B6B6B"
        Accent           = "#B25A1A"
        AccentHover      = "#8F4614"
        AccentPressed    = "#6E3610"
        DataGridAltRow   = "#F9F4EA"
        ButtonHover      = "#EAE3D4"
        ButtonPressed    = "#DDD3BE"
        HeaderBackground = "#EFE8D8"
    }
}

function Get-SavedTheme {
    if (Test-Path $ThemePath) {
        try {
            $data = Get-Content $ThemePath -Raw | ConvertFrom-Json
            if ($data.Theme -in @("Imperial","Rebel")) { return $data.Theme }
        } catch {
            Write-Log "Failed to read theme preference: $($_.Exception.Message)"
        }
    }
    return "Imperial"
}

function Save-Theme {
    param ([string]$ThemeName)
    if (-not (Test-Path $AppDataDir)) { New-Item -ItemType Directory -Path $AppDataDir -Force | Out-Null }
    @{ Theme = $ThemeName } | ConvertTo-Json | Set-Content -Path $ThemePath -Encoding UTF8
}

###############################################################################
# WPF UI XAML
###############################################################################

[xml]$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AT-ST | Assignment and Targeted Scope Tracker"
        Height="820"
        Width="1240"
        WindowStartupLocation="CenterScreen"
        Background="{DynamicResource BackgroundBrush}"
        FontFamily="Segoe UI Variable, Segoe UI"
        FontSize="14">

    <Window.Resources>
        <SolidColorBrush x:Key="BackgroundBrush" Color="#0A0A0C"/>
        <SolidColorBrush x:Key="SurfaceBrush" Color="#161619"/>
        <SolidColorBrush x:Key="SubtleBrush" Color="#1F1F24"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#2E2E35"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#E8E8EB"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#8A8A93"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#B01D22"/>
        <SolidColorBrush x:Key="AccentHoverBrush" Color="#8B1519"/>
        <SolidColorBrush x:Key="AccentPressedBrush" Color="#6B0F13"/>
        <SolidColorBrush x:Key="DataGridAltRowBrush" Color="#131317"/>
        <SolidColorBrush x:Key="ButtonHoverBrush" Color="#1F1F24"/>
        <SolidColorBrush x:Key="ButtonPressedBrush" Color="#2A2A30"/>
        <SolidColorBrush x:Key="HeaderBackgroundBrush" Color="#111114"/>

        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource SurfaceBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="18,9"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontFamily" Value="Segoe UI Variable, Segoe UI"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource ButtonHoverBrush}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource ButtonPressedBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="AccentButton" TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource AccentBrush}"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="{DynamicResource AccentBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="20,9"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontFamily" Value="Segoe UI Variable, Segoe UI"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource AccentHoverBrush}"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="{DynamicResource AccentHoverBrush}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{DynamicResource AccentPressedBrush}"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="{DynamicResource AccentPressedBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="{DynamicResource SurfaceBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Height" Value="36"/>
            <Setter Property="FontFamily" Value="Segoe UI Variable, Segoe UI"/>
        </Style>

        <Style TargetType="ProgressBar">
            <Setter Property="Foreground" Value="{DynamicResource AccentBrush}"/>
            <Setter Property="Background" Value="{DynamicResource SubtleBrush}"/>
            <Setter Property="Height" Value="4"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>

        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="{DynamicResource SurfaceBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="RowBackground" Value="{DynamicResource SurfaceBrush}"/>
            <Setter Property="AlternatingRowBackground" Value="{DynamicResource DataGridAltRowBrush}"/>
            <Setter Property="GridLinesVisibility" Value="None"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="RowHeight" Value="36"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{DynamicResource HeaderBackgroundBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Padding" Value="14,12"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="0,0,0,2"/>
            <Setter Property="FontFamily" Value="Cascadia Mono, Consolas, Segoe UI"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>

        <Style TargetType="DataGridCell">
            <Setter Property="Padding" Value="14,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border Padding="{TemplateBinding Padding}" Background="{TemplateBinding Background}">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="DataGridRow">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{DynamicResource SubtleBrush}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="{DynamicResource SurfaceBrush}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="20"/>
        </Style>

        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{DynamicResource TextPrimary}"/>
            <Setter Property="Margin" Value="0,0,18,0"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
    </Window.Resources>

    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Grid Grid.Row="0" Margin="0,0,0,20">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <Image Grid.Column="0"
                   Name="LogoImage"
                   Width="72"
                   Height="72"
                   VerticalAlignment="Center"
                   Margin="0,0,20,0"/>

            <StackPanel Grid.Column="1" VerticalAlignment="Center">
                <TextBlock Text="AT-ST"
                           FontSize="28"
                           FontWeight="Bold"
                           FontFamily="Cascadia Mono, Consolas, Segoe UI"
                           Foreground="{DynamicResource TextPrimary}"/>
                <TextBlock Text="Assignment and Targeted Scope Tracker"
                           FontSize="13"
                           Foreground="{DynamicResource TextSecondary}"
                           Margin="0,2,0,0"/>
            </StackPanel>

            <Button Grid.Column="2"
                    Name="ThemeToggleButton"
                    Content="Switch to Rebel"
                    MinWidth="150"
                    Height="36"
                    VerticalAlignment="Center"/>
        </Grid>

        <!-- Controls Card -->
        <Border Grid.Row="1" Style="{StaticResource Card}" Margin="0,0,0,16">
            <StackPanel>
                <StackPanel Orientation="Horizontal">
                    <ComboBox Name="GroupDropdown"
                              Width="480"
                              IsEditable="True"
                              IsTextSearchEnabled="False"
                              StaysOpenOnEdit="True"
                              DisplayMemberPath="displayName"
                              SelectedValuePath="id"
                              VirtualizingStackPanel.IsVirtualizing="True"
                              VirtualizingStackPanel.VirtualizationMode="Recycling"
                              ScrollViewer.CanContentScroll="True"
                              ToolTip="Type to filter groups">
                        <ComboBox.ItemsPanel>
                            <ItemsPanelTemplate>
                                <VirtualizingStackPanel/>
                            </ItemsPanelTemplate>
                        </ComboBox.ItemsPanel>
                    </ComboBox>

                    <Button Name="ReloadButton"
                            Content="Reload Groups"
                            Margin="12,0,0,0"/>

                    <Button Name="SearchButton"
                            Content="Run Audit"
                            Style="{StaticResource AccentButton}"
                            Margin="12,0,0,0"/>
                </StackPanel>

                <Grid Margin="0,16,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <TextBlock Name="StatusText"
                               Grid.Column="0"
                               Text="Initializing..."
                               Foreground="{DynamicResource TextSecondary}"
                               VerticalAlignment="Center"/>

                    <TextBlock Name="UserText"
                               Grid.Column="1"
                               FontWeight="SemiBold"
                               Foreground="{DynamicResource AccentBrush}"
                               VerticalAlignment="Center"/>
                </Grid>

                <ProgressBar Name="ProgressBar" Margin="0,12,0,0"/>
            </StackPanel>
        </Border>

        <!-- Results Card (filter bar + grid) -->
        <Border Grid.Row="2" Style="{StaticResource Card}" Padding="0">
            <DockPanel>
                <Border DockPanel.Dock="Top"
                        Background="{DynamicResource HeaderBackgroundBrush}"
                        BorderBrush="{DynamicResource BorderBrush}"
                        BorderThickness="0,0,0,1"
                        CornerRadius="8,8,0,0"
                        Padding="14,10">
                    <DockPanel>
                        <TextBlock Name="ResultCountText"
                                   DockPanel.Dock="Right"
                                   VerticalAlignment="Center"
                                   FontSize="12"
                                   Foreground="{DynamicResource TextSecondary}"/>
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="SHOW:"
                                       FontFamily="Cascadia Mono, Consolas, Segoe UI"
                                       FontWeight="Bold"
                                       FontSize="12"
                                       VerticalAlignment="Center"
                                       Margin="0,0,14,0"/>
                            <CheckBox Name="ShowDirectCheck"     Content="Direct"      IsChecked="True" ToolTip="Assigned directly to the selected group"/>
                            <CheckBox Name="ShowInheritedCheck"  Content="Inherited"   IsChecked="True" ToolTip="Assigned to a parent group the selected group is nested in"/>
                        </StackPanel>
                    </DockPanel>
                </Border>

                <DataGrid Name="ResultsGrid"
                          AutoGenerateColumns="False"
                          CanUserAddRows="False"
                          IsReadOnly="True"
                          EnableRowVirtualization="True"
                          EnableColumnVirtualization="True">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="WORKLOAD" Binding="{Binding Workload}" Width="190"/>
                        <DataGridTextColumn Header="NAME" Binding="{Binding Name}" Width="*"/>
                        <DataGridTextColumn Header="ASSIGNMENT" Binding="{Binding AssignmentType}" Width="110"/>
                        <DataGridTextColumn Header="INTENT" Binding="{Binding Intent}" Width="100"/>
                        <DataGridTextColumn Header="VIA" Binding="{Binding Via}" Width="230"/>
                        <DataGridTextColumn Header="SOURCE" Binding="{Binding Source}" Width="220"/>
                    </DataGrid.Columns>
                </DataGrid>
            </DockPanel>
        </Border>

        <!-- Footer (version) -->
        <TextBlock Grid.Row="3"
                   Name="VersionText"
                   HorizontalAlignment="Right"
                   Margin="0,8,2,0"
                   FontSize="11"
                   Foreground="{DynamicResource TextSecondary}"/>
    </Grid>
</Window>
"@

$Reader = New-Object System.Xml.XmlNodeReader $Xaml
$Window = [Windows.Markup.XamlReader]::Load($Reader)

$GroupDropdown       = $Window.FindName("GroupDropdown")
$ReloadButton        = $Window.FindName("ReloadButton")
$SearchButton        = $Window.FindName("SearchButton")
$StatusText          = $Window.FindName("StatusText")
$UserText            = $Window.FindName("UserText")
$ProgressBar         = $Window.FindName("ProgressBar")
$ResultsGrid         = $Window.FindName("ResultsGrid")
$LogoImage           = $Window.FindName("LogoImage")
$ThemeToggleButton   = $Window.FindName("ThemeToggleButton")
$VersionText         = $Window.FindName("VersionText")

$VersionText.Text = "v$AppVersion"

$ResultCountText     = $Window.FindName("ResultCountText")
$ShowDirectCheck     = $Window.FindName("ShowDirectCheck")
$ShowInheritedCheck  = $Window.FindName("ShowInheritedCheck")

###############################################################################
# RESULTS COLLECTION + VIEW FILTER (Direct / Inherited / All Users / All Devices)
# The background audit adds every match; the view decides what is shown, so
# toggling a checkbox never requires re-running the audit.
###############################################################################

$Global:ResultsCollection = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$Global:ResultsView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($Global:ResultsCollection)

$Global:ShowVia = @{ Direct = $true; Inherited = $true }

$Global:ResultsView.Filter = [Predicate[object]]{
    param($item)
    if ($null -eq $item) { return $false }
    $via = [string]$item.Via
    if ($via -eq "Direct")          { return [bool]$Global:ShowVia.Direct }
    if ($via -like "Inherited*")    { return [bool]$Global:ShowVia.Inherited }
    return $true
}

$ResultsGrid.ItemsSource = $Global:ResultsView

function Update-ResultCount {
    $total = $Global:ResultsCollection.Count
    if ($total -eq 0) { $ResultCountText.Text = ""; return }
    $shown = 0
    foreach ($r in $Global:ResultsView) { $shown++ }
    $ResultCountText.Text = "Showing $shown of $total"
}

$Global:ResultsCollection.Add_CollectionChanged({ Update-ResultCount })

$filterHandler = {
    $Global:ShowVia.Direct     = [bool]$ShowDirectCheck.IsChecked
    $Global:ShowVia.Inherited  = [bool]$ShowInheritedCheck.IsChecked
    $Global:ResultsView.Refresh()
    Update-ResultCount
}
foreach ($cb in @($ShowDirectCheck, $ShowInheritedCheck)) {
    $cb.Add_Checked($filterHandler)
    $cb.Add_Unchecked($filterHandler)
}

###############################################################################
# EMBEDDED LOGO (Base64 PNG)
#
# >>> PASTE YOUR AT-ST LOGO BASE64 BETWEEN THE @' AND '@ MARKERS BELOW <<<
###############################################################################

$LogoBase64 = @'
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
VAjTsmDBApx66qkdAPY9oOCVw3u/9rWvlX71q18d8I6OP/54fOtb38JHP/rR0ymlDxeLvWl88UvMYasAisXeExctWnTcCSecMOM2
v/nNb6CU+knL9NStexsGHIqhwcVi70eLxd6wWOxV+/jaUCz27mYelUoDO6WUb/rkJz+57Z/+6Z9Qq9UOuH3ve9/78OEPf3gRgO8f
8M5S9onD2Ql47aWXXopmae1pafb2P0z+31yrfv9DDz101kzOwDe96U34+7//+4uKxd7OUmlgdJ7bvL+8+d///d/5GWecsU8/uvzy
y9e++OKLKwA82/p5qTSwvljsPeWmm276h1/84hfvf/vb387e+ta3YtmyZfvdwD/+4z/G7bff/nrEATu/2e8dpewTh6UF0HQ6XTNb
6O805n/CD2YbBmSzWZx77rkagKvno62HKqXSwGipNPAnlUplzbe+9a0vXXrppQOXXHIJvvCFL+CXv/wlZpsunA5CCK677joA+OOD
0uCUaTlcLYAL1q1bt2Dp0qUzbjCN+Z/wo7vvvvurn/nMZyjn01++Sy+9FLfddtt1AP51/pp8aFIqDWwG8D+br2N37Njxxh/84Aev
A3B6V1fX4pNPPhmvfe1rce6556KtrW3WfZ1++ukA8IaD3eaUSQ5LCwBzWPd/++23A8B7p46LAfRXq1X68MMPz/jb1772teju7n5N
sdh71Ly2+hCnVBrYUCoNfK1UGriqVBpYMjIysuz222//o7/7u7+7+bzzznP/8z//c9bfd3d3I5/PdxeLvW0vTYtTDjsLoFjszWua
dvkFF8yew+PGG2/c72NQSnHxxRfj29/+9nUAPr3fO3qFUyoN7ABwI4Abi8Xenq9+9av3HnvssUeddtppM/6mra0N1Wq1E0D5JWrm
Yc3haAFcefbZZ2f2N4/dXGn6F95dLPa+ajIvF4u9f1Ys9g4Xi73/u1jsXbwvvy2VBoYAfO+xxx6bdbvmDEu6KOgl4rCzALCPab/2
l6OOOgqrV69e9txzz70ecdrrl5PbP/jBD56Lfb/fmwC82PL/Sz7zmc90bdu27eM//vGPPwL0/gjATQDuKJUGgjnsb83eFO/Y2BgA
DO9jO1P2k8NKARSLvcva2tpef9ZZZ70kx7v00kvx3HPPXYeXWQGUSgNfAfCV+djXwoULccUVV+CDH/ygdtttt139q1/96uonnnii
XCz23gHgYQCPAugDMArAQVwg5DgAH1i+fPlVswVejY2NwfO8sVJpwJ2PtqbsncNKAQC49s1vfjOZyXs/31x00UX48pe/fEWx2Ptn
r7aH2rZtvOMd78A73vEODA8Ptz300ENXbdiw4apNmzZheHgYlUoFruvCtm2sWLECZ599Nq655hpks9kZ99kMs777JTuJlMNOAbx7
th5ovunq6sLpp5+ev//++y8HcMtLduCXmO7ublx66aU40KFVM77itnlpVMqcOGwUQLHYe9ry5ctXz7aUdXBwEG9+85vnvNoNiNOB
feELX5j1+/vvv/86vIoVwHywYcMG3H333TsQLxNOeYk4nGYBZl33DwA/+clPIKX8eqk0QObyAlC84447wtmy5LzxjW9ENps9v1js
nT6hYArK5TI+/elPA8Bfl0oD/svdnsOJw0IBFIu9GiHkHXEpv+mRUuInP/kJMHtm390olQaGPc/7+a9//esZtzEMA+eddx4D8M65
t/iQJdq1a9e87rC/vx8f/OAH8cILL3y+VBr43rzuPGWvHBYKAMDFp5xySufChQtn3OChhx5Cf3//46XSwBP7uO9v/fd///esGxyK
KwT3k699/vOff+bNb34zvvKVr+DJJ5+cMU/i3qhUKrjhhhvw1re+1d20adPHS6WBz8xzW1PmwOHiA7hqbzn/f/zjHwP70Pu38Oun
n366f+vWrQtnSit+yimnoLe39wQAq0qlgS37cYxDglJp4A4AawAcd8MNN1x1ww03XGSa5gknnHACPfbYY7F8+XIsW7YM3d3dyGaz
yGazUEqh0WigXq+jr68Pzz33HJ588kncd999bhiGPwbwd6/ka/JK51UTpTYbxWLvXwP4Ama3eJ4D8JpSaaC6H/v/MwBfBqDNstk2
AKeVSgOlvezr3YgV0cxpiuJ59pNLpYGXPWCmWOwtAHgdgJMAHAVgFYBeADaAHAAFoNp8bQfwFIA/APj5/lzraY7/irpeKSkpKSkp
KSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkHi8MiEOhg8ZMlR6iXuw0pMW/peyF9lveDw2UtwLyTCv+hRXo/9o9UAaSkHMakCmA/
SU3OQ4v0fqSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpLSSTp2kzDuWZa8E8BbE6cD+23HqL87+i8MXy7IvBLAOcSm1Gxynvn9ZVveT
/UoKaln2MQAuQRxH8Ijj1O/Zz/2cjziX3Djik59LgcmXHcuy3wzgBOyl3ZZlZwCcA2AlAB3AZgB3O07dmWbbqwEsn+WwDoDvOk59
7EDaPhfmen6zsOz88y+6XEop77zz149j9wKjsCwbiOsFngygA4AA8AKABxynPnqAzT9gXuJ7ceTHP/6pT9144ze/WSoNteMlLoy6
v1mBT3/Pe67/0wULenu/9KW//zvLsh90nLq3H/tZ+dGPfvKTP/jBTTf19+/sADC4n+15qVn5sY/99V/fcst3vjNTu5sP+VtPO+2M
q1/3unPOAYD77//d7x566P4ey7L/03HqrdtyzvmRn/705z8/0wF//OMf/ODpp5+4C8BBVwCYw/kdIG9YtGjx288776KLensXLnQc
x9mw4ckn77vvt8cD9jccp/6yJe88BO/FQWWfFYBl2cW2tvbVixcvWVIuj4+vXn3McRs3rj8awJOWZecBnI1YexoAdgG4B7HGnO7z
wr/8yz/9U3PXNcuyTwVwCuKMsqz5eQDgvwEcjdhUAgAPsfVRAvBbx6nvbDGlAMBvfr8LwBCAIxFXqR0HcFdiklqWfQSAMwAsABAB
2Iq4OOXrZzoWgC7E2W4T3mNZtgfgFseptwpKe3t7+7rzzrvwwhtu+MY3xsZGRs45501vsixrmeM4GQATxUIdpx5Zlv38Zz/713+d
fPaJT/zN39x66803v/DC8883PwoBXGJZ9qWITesygMcRV+NtPXeveY03IRbcaa+n49RfnOH8SzOdHwBruuvlOPWGZdmnNI+VaznW
HliWbXDOT3/nO9/znkceeeCBW2658UZdN4zjjlu37phjjl331FOPLwUwbFn2kQBOB9DTPPenAdzrOHVhWfZxAF6D2HpwATwL4F4A
52L3Z4AB6Ed8z4eax5/2njtOvbEP9yICcGnzXiTXewTxs0Gm3h/HqWOm85nm+rwewGmIn7kHED/Dc3lGDQCbHKf+85mu/XTsjwVw
3Lp1J530/PObN+/a1de3bt1JJ23cuP5Yy7LXA7jqnHPedPUpp5x2GmOUbty4fv3tt9+2LAzD8XPOedOV03xe++hHP/mFpgXw2mKx
55LLL7/iiq6u7u4dO7ZtcxzHqdWq1aYZeeTHP/6pTz3wwL33nnjiKafk84XCli3PPfeTn/xwoWXZX0u+f+yxhx9es+a449ra2tvX
r3/qqYcfvv/+yy5729uKxQULtm174YVbb725F8BXASyybftdF1542WVHHLFyZRAEwUMP3X//gw/+fgGAzEzHiqLohauuete78vl8
/vrrP/xhAPjlL3/2s0ceefA+7N5TMsY4l1LK4eGh/iAI/vtXv/r5FsSpw/ewlhynPlE70LLsj7R8dYvj1LdZlv2e66//sy8sXLho
URRF0dDQwMAvfvGTnwwM9NPk3B988L771q076STbzuX++Z8//7fFYs8bZ7qelmWbM5z/3TOc32bbts+Y7npZlv1Ud3fx6re85Yor
urqKxdZjTfP85Nra2juz2Wz2/vvvvTeKou96nle+//57XwtgNYABy7KPy+Xy17z5zRdfvGLFypWe53m//vUvfrF587OWZdlDHR2d
77jkkre8ZdGiJUtqtWr1nnvuvHPjxqeLADo//vFPfeqhh+6//4QTTjopm7XtTZvWr//lL3+20LLs/wCweKZ7bln2xPh7Dvfizz/+
8U/97ZTr/bnrr/+zP5/u/liW7cx0PlOuzbmLFy8566qr3vWun/70Rz/aunXLVtu2L5jtGZ3Shk/PLLbTs08KwLJsCmDtunUnnnjP
Pb/5zc6dfX1veMN559m2fXS9Xj9y4cJF604++TWvueGG//iPRqNeX7v2uOOWLl2+ynUdZ7rPt27d8lSyb0rpCVdffe21Tz/9xBM3
3vjNb65YsXLlVVe9610PPHDffa1tWLXq6KO/+91vfUtKpa677v3vP/74E17z+ON/WJN8v3Tp8uU33vjNb+q6Ybz//X/6p0uXLlt2
660331yplMtXX33ddaecctrp999/7+8IISe+4x3XXjswsGvX1772pS9ls7Z9zTXXXTc6OjKyefOzT81yrOd+8IObbmoxkXciFuip
NbP4+PjYWKNRr1955Tuv/f73vxtEUfQ9xL6Av7QsmwD4mePU51oU44VvfOP//n8AiKZp2gknnHzylVe+853/9/9++QXZrGZ69NFr
137ve9/+drlcLlNK2SzXkxBCzp/p/Kc7P0LI6pm2f/75zYWrr7722g0bnn76xhv/8z9Xrly16oorrrlm6r1rUqlUymP1er1+7rkX
XPDYY4+0j46O9CultiIuDDpGCHnHNddce+3Q0ODg17/+L/+iabp+2mlnnLF587NLKKXHXHPNdddt3vzssz/4wU03LV68dOlVV73r
XWNjIyMDA/0vAMCxxx5//E033XCDUsDVV1977Zlnnn3+vffevZ4QcvQs9/xRAE/O8V5g6vUGMDLT/VFKkZnOp3V/K1asPPetb73q
qltv/d73duzYfhshZOXentEpbXh+j0buhX1dDHTE0qXL1+Ry+fxzzz2zcXx8bNvgYH//ccedcAKAI4QQQogo8n3f9zyv9Nhjjz66
deuWLTN9jtj5AwDo7V10hGGY5r333vMr3/f/97PPbrp9y5bnnpvagN/97q67KpXKf9Vq1TuefXbjxmJxwQIA7ZPf3313rVb79ujo
yMM7dmzbtmHDU08NDg78zHXdXz7zzIYNPT0LFgDo7elZcER3d7H461//4heu67ojI8PDDz10//1r1hx3HJom7AzHqgOotTTp245T
/5cp5j8AXPTWt1555eOPP/pouTw+fvXV1/4x5/xdAM6/7LK3X3Pmma8/F7FZN1f+AOARAOUwDMNHH33oIUIo7ejoXILmbM7vfnfX
XeVy+TsAvtPbu2jpLNfT3sv5h63nB+DWnp4Fy2favrd30YpMJpP53e/uus33/S9t2rTh9uef37x5upNwnHoYhuFd3/72N75h27nc
e9/7J3/yiU98+nOXXvq2/2VZ1ocArFuwoHdFZ2d39223/fTWer3+f8bHx77y61//4kYApd7ehctzuXz+7rvv+IXneV96/vnNv9y4
cf365jM4cd/Gx8dvLpfHb/7tb++887jj1q0DsGIv5zxz3bgZSK6349T/EbHymvb+zHI+25J9rV69Zs1b3nLllTff/F//tWPH9h8B
eGquz2jShn01/4F9HwIcd8IJJ520ZctzzwVBEAHIb9jw9NPr1p100oMP/n750NDgY/fcc+fCK6+85hpKKd2yZfPmBx+87+6hocGn
7rnnzjunfh6G4cQ0ZC6Xy9Vq1aqUctxx6o5l2cOVSrk8tQGO02ggnjLpCMMoymQsCy0VeZrflwEEURRFjuM4zf/LKIoizjUNgFEo
tLWZpml++tOf/zwhhABxt/jii1u3oilQezvWTFiW3dnR0XnM0UevXfuTn9x6qxDCv/zyt7/jmmuu++Pvf/+73122bPnyRx998EHE
FWv2StNauPrUU1972Wtfe+aZbW3t7YwxBgC2ncuNjAwTAGia3CUA7Xu5ntpezn9qx5CdbftcLperVieO1bAse6hcHh+f6Xwcp/7Q
yAhGfvSjWx4EsLijo3PRpZe+7W2XXfb2K2+55Ts78/lCoVIpl6MoGnGcerKfb1iWvSqXy7+lWq1UhBBjjlOvW5bdPz4+NlYsFoto
3rfmeQ4CUOXy+Hgul8/v7RwAFOZyL1pJrvfe7o9hGMZ059O8t6sA4LWvPeOMjRvXrx8Y6H8ewO8BLJvLM9pyz/eLOSsAy7JNTdOO
Wbv2+OMNwzA++9kvfrH1+wULeo8eHBzoe+qpxx9/6qnHH9B1Pf+Wt1x55dlnv/GCO++8/VczfP6z5Pf1eq2Wy+XzlNL25vRZd6HQ
1jY8XNrvk5sFUa1WKrVatfrlL//jPyql9jWZhGr5zXSxFLqmaRoAEEKkUupnP/vZj7W3v/0d7/jgB//iL0qloaGBgf5diEtlzYUF
XV3da88994ILvvOdb31raGhgIIqi6KMf/cQnKKUTwtrSpvperme0t/Ofcn7ubNsvWrRkyZRjFdva2ttLpaGhqfu1LFsH8D7EVtTv
Afx0bGz0NY8++mDvRRddfjkAu1qtVAqFtjbOeZdl2W0AMgAuBTBaq1Wr+XyhwBjrsCw7C2Bhe3tHR7VarSJ2vqFQaGvr79/VA0C1
tbW3N4XEOcB7vgct+5j1/sxyPjuSfd166y23XHTRZZedddY5F91332+rAH4/l/Ye6HnsiwWw5uij1xwXRVH0z//8+c8LISbM92uv
fd/71q076aRC4YW27u5i8amnnngiDIOAEEI0TddXrz5m7XSfo0V4+vt3veD7nvf617/hwgcf/H1u+fIjjli1avXqg6QA3MHBgRcr
lXL5/PMvuuj3v//d75RSavXqY47RNE175JEH9/DOTv19o1Gv9/T09AwM7HqPZdkudp8FGBkeLu0aHx8bu/zyK95xxx232VEURbt2
7dy5Zs1xx42MDA8TQqhSKg+gMof2Kk3TNCml9DzXpZTSM84466y2tvb2GbYf7u/ftXWW61nfy/nf2Xp+Sil3cHBg20zb/+EPD98f
H+uNFz344H25FStWrjzyyKOOmk4BACh0dxePOuusc855+OEHTi+VhoZyuXz+1FNf+9qmv6EyODjw4ujo8PDFF19+xV133Z7TNE07
7bQzz/z1r3/x3YGB/m31eq32xjeef8l9992TXbx46dJjjz3++G9969//Hc0h5etf/8Y3Dgz09wNKnX32uedu2PD00wBeHBwcMGc5
59/P4T7s1/0ZHBzYPtP5JDtoNOr1G2/85jff857rrxdCiAceuA8H+IzOiX1RAMesW3fSSU8++dhjQojHHaf+KwCwLPuExx575NiL
Lrr88t/+9s47Fy9esuT66z/8YU3T9Rdf3Lr1d7+7++4wDJzpPkfLWFNK+dQPfnDTTZdd9ra3nX76WWft2LFt2zPPbNjQVDRipkbt
J0opdfstt3yn4/zzL77kQx/6yEcA4Lnnnn32zjt//cs5HO+F++67554LL7zssssue/vbf/Wrn/+8dRbAceqhZdl333TTt7suuODi
i//0Tz/6UUopff755577t3/76levuOKaa84//6KLbr/9tn4A/zWH9pYGBvo3PPnkY4994AMf/nAYhuFzz23aNIOAJce/f5brGe7l
/LdNc37fuOWW79w43fZSyidvueU737n88iuuOP30171ux45t29avf+qp6doGYGxkZHjLli3P9V544aWX9vT09vq+573wwtatd9zx
y18CuF8pZdxyy3e6L7zwsss+/OG//EvHcZzbb7/tNgB9Usrhm2++ccEll7zlLR/72F//dbVaqfzsZz/60cBA/yNo+oI2bVq//t3v
fu97bTuX27Rp/fr77//dbwA8oJTaMcs575ihvXNh1vujlHr8llu+853pzgfAEclOarXaXf/1X9+k73nPBz4gpZQPPXT/j2e65pgn
mZhzKHBL1J6HOApqpPm5DuA6xHOgJcRj5uSkBgE8iLiXe/00n68GcCLi+fn7AFzWesxrr33f+zZtWr/+scce/QLiirPHN/d/A2Kn
zRWIx6u/aO679fu1AM5HrGRuQWweXoPYC3+H49SftCx7OeI51t7mIXcinus/eS/H2grgPMQRftnmNZkaB4Dm/s8E0N387UDzXI8G
YALY1jrl1PK7SwAci7ii7rebY/h2xPPcidNPb25ebbZ7DVrCSS3LPgrA22a6no5Tf2KW8x+b7vyabd5je8epD1mW/ZrmdWuNOQgB
/HhqKLBl2QbimIE1iOMzPMTPxEOOU+9rbrMK8bx5EXHswnpMxgGsQxwH0I44DuC5Zrv/5OMf/9Q/NKPqhhALyQDiOICBlnsy7TlM
vQ+z3IuL0PJ8NK/3bPfn2wAWTXc+AC7Ans/tuc3fP4t4ZmKvz+hLHUI871iW/cbXv/7c/z766GP/vb2982/POOPsH/793385WLBg
4T9Ylt35crfvlcbheD0ty/7zz3zmi7Xly1d+xbLs7pe7Pa8E9jcU+GCw84UXnn/+0kvf9rZFixYvHh8fG7vllu98p1qtHhLx4a9A
0uuZslcOqdWAlmUfjdhk7kA8374RwP2OU59vH8BhweF2PaczzV/mJqWkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSk
pKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSk
pKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKSkpKS8svj/AUuaiIiqbUHaAAAAAElFTkSuQmCC
'@

function Load-Logo {
    try {
        $cleanBase64 = $LogoBase64 -replace '\s',''
        if ([string]::IsNullOrWhiteSpace($cleanBase64) -or $cleanBase64 -eq "PASTE_YOUR_BASE64_HERE") {
            Write-Log "No logo Base64 embedded; skipping logo load."
            return
        }
        $bytes = [System.Convert]::FromBase64String($cleanBase64)
        $stream = New-Object System.IO.MemoryStream(,$bytes)
        $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
        $bitmap.BeginInit()
        $bitmap.StreamSource = $stream
        $bitmap.CacheOption = 'OnLoad'
        $bitmap.EndInit()
        $bitmap.Freeze()
        $LogoImage.Source = $bitmap
    } catch {
        Write-Log "Failed to load embedded logo: $($_.Exception.Message)"
    }
}

Load-Logo

# Window/taskbar icon. The window belongs to pwsh.exe (started by the
# launcher), so it needs its own icon or it shows the PowerShell icon.
$IconPath = Join-Path $AppRoot "AT-ST.ico"
try {
    if (Test-Path $IconPath) {
        $Window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$IconPath)
    } elseif ($LogoImage.Source) {
        $Window.Icon = $LogoImage.Source
    }
} catch {
    Write-Log "Failed to set window icon: $($_.Exception.Message)"
}

###############################################################################
# THEME APPLICATION
###############################################################################

$Global:CurrentTheme = Get-SavedTheme

function Apply-Theme {
    param ([string]$ThemeName)

    if (-not $Themes.ContainsKey($ThemeName)) { $ThemeName = "Imperial" }
    $palette = $Themes[$ThemeName]

    $keyToPaletteName = @{
        "BackgroundBrush"       = "Background"
        "SurfaceBrush"          = "Surface"
        "SubtleBrush"           = "Subtle"
        "BorderBrush"           = "Border"
        "TextPrimary"           = "TextPrimary"
        "TextSecondary"         = "TextSecondary"
        "AccentBrush"           = "Accent"
        "AccentHoverBrush"      = "AccentHover"
        "AccentPressedBrush"    = "AccentPressed"
        "DataGridAltRowBrush"   = "DataGridAltRow"
        "ButtonHoverBrush"      = "ButtonHover"
        "ButtonPressedBrush"    = "ButtonPressed"
        "HeaderBackgroundBrush" = "HeaderBackground"
    }

    foreach ($key in $keyToPaletteName.Keys) {
        $paletteName = $keyToPaletteName[$key]
        $hex = $palette[$paletteName]

        [System.Windows.Media.Color]$colorLocal = [System.Windows.Media.ColorConverter]::ConvertFromString($hex)
        [System.Windows.Media.SolidColorBrush]$brushLocal = [System.Windows.Media.SolidColorBrush]::new($colorLocal)
        $brushLocal.Freeze()

        if ($Window.Resources.Contains($key)) {
            [void]$Window.Resources.Remove($key)
        }
        [void]$Window.Resources.Add($key, $brushLocal)
    }

    if ($ThemeName -eq "Imperial") {
        $ThemeToggleButton.Content = "Switch to Rebel"
    } else {
        $ThemeToggleButton.Content = "Switch to Imperial"
    }

    $Global:CurrentTheme = $ThemeName
    Save-Theme -ThemeName $ThemeName
    Write-Log "Applied theme: $ThemeName"
}

Apply-Theme -ThemeName $Global:CurrentTheme

$ThemeToggleButton.Add_Click({
    if ($Global:CurrentTheme -eq "Imperial") {
        Apply-Theme -ThemeName "Rebel"
    } else {
        Apply-Theme -ThemeName "Imperial"
    }
})

###############################################################################
# UI HELPERS (thread-safe)
###############################################################################

function Invoke-OnUI {
    param ([scriptblock]$Action)
    $Window.Dispatcher.Invoke($Action)
}

function Set-Status {
    param (
        [string]$Text,
        [Nullable[double]]$Progress = $null
    )
    Invoke-OnUI {
        $StatusText.Text = $Text
        if ($Progress -ne $null) {
            if ($Progress -lt 0) { $Progress = 0 }
            if ($Progress -gt 100) { $Progress = 100 }
            $ProgressBar.Value = $Progress
        }
    }
}

function Add-Result {
    param ($ResultObject)
    Invoke-OnUI { $Global:ResultsCollection.Add($ResultObject) }
}

function Clear-Results {
    Invoke-OnUI { $Global:ResultsCollection.Clear() }
}

###############################################################################
# GRAPH HELPERS (main thread)
###############################################################################

# Shared helpers as source text so they can be loaded on the UI thread, in the
# audit runspace, and inside ForEach-Object -Parallel workers.
$GraphHelpersSrc = @'
function Invoke-GraphWithRetry {
    param ([string]$Uri, [int]$MaxAttempts = 5)
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return Invoke-MgGraphRequest -Uri $Uri -Method GET -ErrorAction Stop
        } catch {
            $status = $null
            $resp = $_.Exception.Response
            try { if ($resp -and $resp.StatusCode) { $status = [int]$resp.StatusCode } } catch { }
            if (-not $status -and $_.Exception.Message -match '\b(429|503|504)\b') { $status = [int]$Matches[1] }

            if (($status -notin @(429, 503, 504)) -or $attempt -ge $MaxAttempts) { throw }

            $wait = $null
            try {
                if ($resp.Headers.RetryAfter.Delta) { $wait = [double]$resp.Headers.RetryAfter.Delta.TotalSeconds }
            } catch { }
            if (-not $wait -or $wait -lt 1) { $wait = [math]::Pow(2, $attempt) }   # 2,4,8,16
            $wait = [math]::Min($wait, 60) + ((Get-Random -Minimum 0 -Maximum 1000) / 1000)
            Start-Sleep -Milliseconds ([int]($wait * 1000))
        }
    }
}

function Get-PagedResults {
    param ([Parameter(Mandatory)][string]$Uri)
    $All = @()
    do {
        $Response = Invoke-GraphWithRetry -Uri $Uri
        if ($Response.value) { $All += $Response.value }
        $Uri = $Response.'@odata.nextLink'
    } while ($Uri)
    return $All
}

function Get-AssignmentMatches {
    param ($Item, $Assignments, $TargetMap, $Spec, $WorkloadName, $IntentMode)
    $out = @()
    foreach ($a in @($Assignments)) {
        if (-not $a -or -not $a.target) { continue }
        $ttype = [string]$a.target.'@odata.type'
        $via = $null
        if ($Spec) {
            if ($Spec -eq "allUsers" -and $ttype -like "*allLicensedUsers*") { $via = "All Users" }
            elseif ($Spec -eq "allDevices" -and $ttype -like "*allDevices*") { $via = "All Devices" }
        }
        elseif ($a.target.groupId -and $TargetMap.ContainsKey([string]$a.target.groupId)) {
            $via = $TargetMap[[string]$a.target.groupId]
        }
        if (-not $via) { continue }

        if ($Item.displayName) { $itemName = $Item.displayName }
        elseif ($Item.name) { $itemName = $Item.name }
        else { $itemName = $Item.id }

        if ($ttype -like "*exclusion*") { $atype = "Exclude" } else { $atype = "Include" }

        switch ($IntentMode) {
            "enforced" { $intent = "Enforced" }
            "none"     { $intent = "" }
            default    { $intent = $a.intent }
        }

        $out += [PSCustomObject]@{
            Workload       = $WorkloadName
            Name           = $itemName
            AssignmentType = $atype
            Intent         = $intent
            Via            = $via
            Source         = $Item.'@odata.type'
        }
    }
    return $out
}
'@

. ([scriptblock]::Create($GraphHelpersSrc))

###############################################################################
# DELEGATED AUTH  (browser interactive; device-code fallback)
###############################################################################

function Test-ScopesSatisfied {
    param ($Context)
    if (-not $Context -or -not $Context.Account) { return $false }
    if ($EnterpriseAppClientId -and $Context.ClientId -and ($Context.ClientId -ne $EnterpriseAppClientId)) { return $false }
    # Log (don't block on) permissions that aren't admin-consented yet
    $missing = @($RequiredScopes | Where-Object { $Context.Scopes -notcontains $_ })
    if ($missing.Count -gt 0) {
        Write-Log "Missing admin-consented scopes (affected workloads will be skipped): $($missing -join ', ')"
    }
    return $true
}

function Try-BrowserSignIn {
    try {
        Write-Log "Attempting browser interactive sign-in"
        # No -Scopes: requests .default (admin-consented permissions only, no user consent prompts)
        Connect-MgGraph -TenantId $TenantId -ClientId $EnterpriseAppClientId `
            -NoWelcome -ErrorAction Stop | Out-Null
        Write-Log "Browser sign-in succeeded"
        return $true
    } catch {
        Write-Log "Browser sign-in failed: $($_.Exception.Message)"
        return $false
    }
}

function Try-DeviceCodeSignIn {
    try {
        Write-Log "Attempting device code sign-in"
        Set-Status -Text "Browser sign-in unavailable. Follow the device code prompt..."
        Connect-MgGraph -TenantId $TenantId -ClientId $EnterpriseAppClientId `
            -UseDeviceCode -NoWelcome -ErrorAction Stop | Out-Null
        Write-Log "Device code sign-in succeeded"
        return $true
    } catch {
        Write-Log "Device code sign-in failed: $($_.Exception.Message)"
        return $false
    }
}

function Ensure-GraphConnection {
    $context = Get-MgContext
    if (Test-ScopesSatisfied -Context $context) {
        $script:SignedInUser = $context.Account
        Invoke-OnUI { $UserText.Text = "Signed in as $($context.Account)" }
        Write-Log "Using cached session for $($context.Account)"
        return
    }

    Set-Status -Text "Signing in to Microsoft Graph..."
    $signedIn = Try-BrowserSignIn
    if (-not $signedIn) { $signedIn = Try-DeviceCodeSignIn }
    if (-not $signedIn) { throw "Microsoft Graph sign-in did not complete successfully." }

    $context = Get-MgContext
    if (-not $context -or -not $context.Account) { throw "Microsoft Graph sign-in did not return a valid session." }
    $script:SignedInUser = $context.Account
    Invoke-OnUI { $UserText.Text = "Signed in as $($context.Account)" }
    Write-Log "Signed in as $($context.Account)"
}

###############################################################################
# GROUP CACHE + LOAD
###############################################################################

function Load-Groups {
    param ([switch]$ForceRefresh)

    if (-not $ForceRefresh -and (Test-Path $CachePath)) {
        try {
            Set-Status -Text "Loading groups from cache..."
            $cached = @(Get-Content -Path $CachePath -Raw | ConvertFrom-Json)
            if ($cached.Count -gt 0) {
                Write-Log "Loaded $($cached.Count) groups from cache"
                return @($cached | Sort-Object displayName)
            }
        } catch {
            Write-Log "Cache load failed; refreshing from Graph"
        }
    }

    Ensure-GraphConnection
    Set-Status -Text "Loading groups from Microsoft Graph..."

    $groups = @(Get-PagedResults "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName")
    $groups = @($groups | Sort-Object displayName)

    if (-not (Test-Path $AppDataDir)) { New-Item -ItemType Directory -Path $AppDataDir -Force | Out-Null }
    $groups | ConvertTo-Json -Depth 5 | Set-Content -Path $CachePath -Encoding UTF8
    Write-Log "Loaded $($groups.Count) groups from Graph"
    return @($groups)
}

function Get-SelectedGroup {
    # Only trust SelectedItem if the text still matches it
    $sel = $GroupDropdown.SelectedItem
    if ($sel -and ([string]$GroupDropdown.Text -eq [string]$sel.displayName)) { return $sel }
    $typed = $GroupDropdown.Text
    if ([string]::IsNullOrWhiteSpace($typed)) { return $null }
    $match = @(@($Global:SpecialGroups) + @($Global:AllGroups) | Where-Object { $_.displayName -eq $typed })
    if ($match.Count -eq 1) { return $match[0] }
    return $null
}

###############################################################################
# COLLECTIONVIEW FILTERING SETUP  (fast dropdown)
###############################################################################

$Global:GroupsBacking = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$Global:GroupsView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($Global:GroupsBacking)
$Global:GroupFilterText = ""

$Global:GroupsView.Filter = [Predicate[object]]{
    param($item)
    if ([string]::IsNullOrWhiteSpace($Global:GroupFilterText)) { return $true }
    if ($null -eq $item) { return $false }
    return ([string]$item.displayName) -like "*$($Global:GroupFilterText)*"
}

$GroupDropdown.ItemsSource = $Global:GroupsView

# Built-in targets shown at the top of the dropdown. Not real groups:
# auditing them matches All Users / All Devices assignment targets.
$Global:SpecialGroups = @(
    [PSCustomObject]@{ id = "__ALLUSERS__";   displayName = "[All Users]" }
    [PSCustomObject]@{ id = "__ALLDEVICES__"; displayName = "[All Devices]" }
)

function Set-GroupsBacking {
    param ($Groups)
    $Global:GroupsBacking.Clear()
    foreach ($g in $Global:SpecialGroups) { [void]$Global:GroupsBacking.Add($g) }
    foreach ($g in $Groups) { [void]$Global:GroupsBacking.Add($g) }
    $Global:GroupFilterText = ""
    $Global:GroupsView.Refresh()
}

###############################################################################
# INITIAL LOAD
###############################################################################

try {
    Ensure-GraphConnection
    $Global:AllGroups = @(Load-Groups)
    Set-GroupsBacking -Groups $Global:AllGroups
    Set-Status -Text "Ready. All systems operational."
} catch {
    Write-Log "Init failed: $($_.Exception.Message)"
    [System.Windows.MessageBox]::Show($_.Exception.Message, "$AppName - Initialization Failed", "OK", "Error") | Out-Null
    Set-Status -Text "Initialization failed"
}

###############################################################################
# COMBOBOX TYPE-TO-FILTER (debounced + CollectionView refresh)
###############################################################################

$GroupDropdown.Add_Loaded({
    $tb = $GroupDropdown.Template.FindName("PART_EditableTextBox", $GroupDropdown)
    if ($tb) {
        $Global:FilterTimer = New-Object System.Windows.Threading.DispatcherTimer
        $Global:FilterTimer.Interval = [TimeSpan]::FromMilliseconds(200)

        $Global:FilterTimer.Add_Tick({
            $Global:FilterTimer.Stop()
            $newText = $GroupDropdown.Text
            if ($null -eq $newText) { $newText = "" }

            # Text was set by selecting a group, not by typing - don't filter/reopen
            $sel = $GroupDropdown.SelectedItem
            if ($sel -and $newText -eq [string]$sel.displayName) { return }

            if ($newText -ne $Global:GroupFilterText) {
                $Global:GroupFilterText = $newText
                $Global:GroupsView.Refresh()
                if (-not [string]::IsNullOrWhiteSpace($newText) -and -not $GroupDropdown.IsDropDownOpen) {
                    $GroupDropdown.IsDropDownOpen = $true
                }
            }
        })

        $Global:GroupTextBox = $tb

        $tb.Add_TextChanged({
            $Global:FilterTimer.Stop()
            $sel = $GroupDropdown.SelectedItem
            if ($sel -and [string]$GroupDropdown.Text -eq [string]$sel.displayName) { return }
            $Global:FilterTimer.Start()
        })
    }
})

# When a group is picked: close the dropdown, clear the highlight, and move
# focus to Run Audit so one click (or Enter) starts the audit.
$GroupDropdown.Add_SelectionChanged({
    if (-not $GroupDropdown.SelectedItem) { return }
    if ($Global:FilterTimer) { $Global:FilterTimer.Stop() }
    $GroupDropdown.IsDropDownOpen = $false

    $Window.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Input, [Action]{
        if ($Global:GroupTextBox) {
            $Global:GroupTextBox.SelectionLength = 0
            $Global:GroupTextBox.CaretIndex = $Global:GroupTextBox.Text.Length
        }
        [void]$SearchButton.Focus()
    }) | Out-Null
})

###############################################################################
# RELOAD GROUPS
###############################################################################

$ReloadButton.Add_Click({
    try {
        $ReloadButton.IsEnabled = $false
        $SearchButton.IsEnabled = $false
        Remove-Item -Path $CachePath -ErrorAction SilentlyContinue
        $Global:AllGroups = @(Load-Groups -ForceRefresh)
        Set-GroupsBacking -Groups $Global:AllGroups
        Set-Status -Text "Groups reloaded"
    } catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, "$AppName - Reload Failed", "OK", "Error") | Out-Null
    } finally {
        $ReloadButton.IsEnabled = $true
        $SearchButton.IsEnabled = $true
    }
})

###############################################################################
# AUDIT ENGINE (background runspace; manifest-driven; parallel per workload)
###############################################################################

function Start-AuditRunspace {
    param (
        [string]$GroupId,
        [string]$GroupName
    )

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()

    $runspace.SessionStateProxy.SetVariable("GroupId", $GroupId)
    $runspace.SessionStateProxy.SetVariable("GroupName", $GroupName)
    $runspace.SessionStateProxy.SetVariable("Window", $Window)
    $runspace.SessionStateProxy.SetVariable("StatusText", $StatusText)
    $runspace.SessionStateProxy.SetVariable("ProgressBar", $ProgressBar)
    $runspace.SessionStateProxy.SetVariable("ResultsCollection", $Global:ResultsCollection)
    $runspace.SessionStateProxy.SetVariable("ParallelThrottleLimit", $ParallelThrottleLimit)
    $runspace.SessionStateProxy.SetVariable("LogPath", $LogPath)
    $runspace.SessionStateProxy.SetVariable("WorkloadManifest", $WorkloadManifest)
    $runspace.SessionStateProxy.SetVariable("GraphHelpersSrc", $GraphHelpersSrc)
    $runspace.SessionStateProxy.SetVariable("GraphModuleRef", $GraphModuleRef)

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $runspace

    $ps.AddScript({

        Import-Module $GraphModuleRef -ErrorAction Stop

        function Invoke-OnUI { param([scriptblock]$Action) $Window.Dispatcher.Invoke($Action) }

        function Set-Status {
            param ([string]$Text, [Nullable[double]]$Progress = $null)
            Invoke-OnUI {
                $StatusText.Text = $Text
                if ($Progress -ne $null) { $ProgressBar.Value = $Progress }
            }
        }

        function Add-Result {
            param ($Obj)
            Invoke-OnUI { $ResultsCollection.Add($Obj) }
        }

        function Write-Log {
            param ([string]$Message)
            $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Add-Content -Path $LogPath -Value "$ts | $Message"
        }

        # Retry-aware Graph helpers (Invoke-GraphWithRetry, Get-PagedResults, Get-AssignmentMatches)
        . ([scriptblock]::Create($GraphHelpersSrc))

        function Add-FlattenedResults {
            param ($Batches)
            foreach ($batch in $Batches) {
                if ($batch) {
                    foreach ($item in @($batch)) {
                        if ($item -and $item.PSObject.Properties['Workload']) { Add-Result $item }
                    }
                }
            }
        }

        try {
            Write-Log "Audit started for group '$GroupName' [$GroupId]"

            # ---- Built-in targets (All Users / All Devices) ----
            $SpecialTarget = $null
            if ($GroupId -eq "__ALLUSERS__")   { $SpecialTarget = "allUsers" }
            if ($GroupId -eq "__ALLDEVICES__") { $SpecialTarget = "allDevices" }

            # ---- Effective target map: selected group + all parent groups ----
            $TargetMap = @{}
            $FailedItems     = [System.Collections.Generic.List[string]]::new()
            $FailedWorkloads = [System.Collections.Generic.List[string]]::new()
            if ($SpecialTarget) {
                Write-Log "Auditing built-in target: $SpecialTarget"
            } else {
            Set-Status -Text "Resolving group memberships..." -Progress 0
            $TargetMap[$GroupId] = "Direct"
            try {
                $parents = @(Get-PagedResults "https://graph.microsoft.com/v1.0/groups/$GroupId/transitiveMemberOf?`$select=id,displayName")
                foreach ($p in $parents) {
                    if ($p.'@odata.type' -eq '#microsoft.graph.group' -and $p.id -and -not $TargetMap.ContainsKey($p.id)) {
                        $TargetMap[$p.id] = "Inherited: $($p.displayName)"
                    }
                }
                Write-Log "Resolved $($TargetMap.Count - 1) parent group(s) for '$GroupName'"
            } catch {
                Write-Log "Parent group lookup failed (direct assignments only): $($_.Exception.Message)"
            }
            }

            $total = @($WorkloadManifest).Count
            $index = 0

            foreach ($wl in $WorkloadManifest) {
                $index++
                $baseProgress = [int](($index - 1) / $total * 100)
                Set-Status -Text "Loading $($wl.Name)..." -Progress $baseProgress

                if ($wl.Type -eq "rbac") {
                    # RBAC: single list + membership/scope check
                    try {
                        $roles = @(Get-PagedResults $wl.ListUri)
                        foreach ($ra in $roles) {
                            if ($SpecialTarget) {
                                # Role assignments can be scoped to all users/devices
                                $st = [string]$ra.scopeType
                                $hit = ($SpecialTarget -eq "allUsers"   -and $st -in @("allLicensedUsers","allDevicesAndLicensedUsers")) -or
                                       ($SpecialTarget -eq "allDevices" -and $st -in @("allDevices","allDevicesAndLicensedUsers"))
                                if ($hit) {
                                    Add-Result ([PSCustomObject]@{
                                        Workload       = $wl.Name
                                        Name           = $ra.displayName
                                        AssignmentType = "Scope"
                                        Intent         = ""
                                        Via            = if ($SpecialTarget -eq "allUsers") { "All Users" } else { "All Devices" }
                                        Source         = "deviceManagement/roleAssignments"
                                    })
                                }
                                continue
                            }
                            foreach ($tid in @($TargetMap.Keys)) {
                                $role = $null
                                if ($ra.members -contains $tid) { $role = "Member" }
                                elseif ($ra.scopeMembers -contains $tid) { $role = "Scope" }
                                if ($role) {
                                    Add-Result ([PSCustomObject]@{
                                        Workload       = $wl.Name
                                        Name           = $ra.displayName
                                        AssignmentType = $role
                                        Intent         = ""
                                        Via            = $TargetMap[$tid]
                                        Source         = "deviceManagement/roleAssignments"
                                    })
                                }
                            }
                        }
                    } catch {
                        Write-Log "$($wl.Name) failed: $($_.Exception.Message)"
                        $FailedWorkloads.Add($wl.Name)
                    }
                    continue
                }

                # assignments-type workload
                $wlName    = $wl.Name
                $wlBaseUri = $wl.ListUri
                $wlIntent  = $wl.IntentMode

                # 1) Try one paged list call with $expand=assignments
                $items = $null
                $expanded = $false
                try {
                    $items = @(Get-PagedResults "$($wlBaseUri)?`$expand=assignments")
                    $withAssignments = @($items | Where-Object { $_ -is [System.Collections.IDictionary] -and $_.Contains('assignments') })
                    if ($items.Count -eq 0 -or $withAssignments.Count -gt 0) {
                        $expanded = $true
                    } else {
                        Write-Log "$($wlName): `$expand=assignments ignored by service; using per-item lookups"
                    }
                } catch {
                    Write-Log "$($wlName): `$expand=assignments not supported ($($_.Exception.Message)); using per-item lookups"
                    $items = $null
                }

                if ($expanded) {
                    Set-Status -Text "Checking $($items.Count) $wlName..." -Progress $baseProgress
                    foreach ($it in $items) {
                        $assigns = if ($it -is [System.Collections.IDictionary]) { $it['assignments'] } else { $it.assignments }
                        foreach ($r in @(Get-AssignmentMatches -Item $it -Assignments $assigns -TargetMap $TargetMap -Spec $SpecialTarget -WorkloadName $wlName -IntentMode $wlIntent)) {
                            if ($r) { Add-Result $r }
                        }
                    }
                    Write-Log "$($wlName): $($items.Count) item(s) checked via `$expand"
                    continue
                }

                # 2) Fallback: list items, then per-item /assignments in parallel
                try {
                    $items = @(Get-PagedResults $wlBaseUri)
                } catch {
                    Write-Log "$($wlName) list failed: $($_.Exception.Message)"
                    $FailedWorkloads.Add($wlName)
                    continue
                }

                Set-Status -Text "Scanning $($items.Count) $wlName..." -Progress $baseProgress

                $batches = $items | ForEach-Object -Parallel {
                    Import-Module $using:GraphModuleRef -ErrorAction SilentlyContinue
                    . ([scriptblock]::Create($using:GraphHelpersSrc))
                    $it = $_
                    try {
                        $resp = Invoke-GraphWithRetry -Uri "$($using:wlBaseUri)/$($it.id)/assignments"
                        Get-AssignmentMatches -Item $it -Assignments $resp.value -TargetMap $using:TargetMap `
                            -Spec $using:SpecialTarget -WorkloadName $using:wlName -IntentMode $using:wlIntent
                    } catch {
                        $n = if ($it.displayName) { $it.displayName } elseif ($it.name) { $it.name } else { $it.id }
                        [PSCustomObject]@{ __Failed = $true; Name = $n; Error = $_.Exception.Message }
                    }
                } -ThrottleLimit $ParallelThrottleLimit

                foreach ($r in @($batches)) {
                    if (-not $r) { continue }
                    if ($r.PSObject.Properties['__Failed']) {
                        $FailedItems.Add("$wlName / $($r.Name): $($r.Error)")
                    } elseif ($r.PSObject.Properties['Workload']) {
                        Add-Result $r
                    }
                }
                Write-Log "$($wlName): $($items.Count) item(s) checked via per-item lookups"
            }

            foreach ($f in $FailedItems) { Write-Log "FAILED ITEM: $f" }
            foreach ($w in $FailedWorkloads) { Write-Log "FAILED WORKLOAD: $w" }

            $gaps = @()
            if ($FailedWorkloads.Count -gt 0) { $gaps += "$($FailedWorkloads.Count) workload(s) could not be read" }
            if ($FailedItems.Count -gt 0)     { $gaps += "$($FailedItems.Count) item(s) could not be checked" }

            if ($gaps.Count -gt 0) {
                Set-Status -Text "Audit complete with gaps: $($gaps -join '; '). See audit.log." -Progress 100
            } else {
                Set-Status -Text "Audit complete. All systems reported." -Progress 100
            }
            Write-Log "Audit complete for '$GroupName'"

        } catch {
            $errMsg = $_.Exception.Message
            Write-Log "Audit failed: $errMsg"
            Invoke-OnUI {
                [System.Windows.MessageBox]::Show($errMsg, "Audit Failed", "OK", "Error") | Out-Null
                $StatusText.Text = "Audit failed"
            }
        }
    }) | Out-Null

    $handle = $ps.BeginInvoke()

    # Completion watcher. Runs on the WPF dispatcher, so it fires while the
    # window is open (Register-ObjectEvent actions do not, which left the
    # buttons disabled after the first audit).
    $Global:AuditPS       = $ps
    $Global:AuditHandle   = $handle
    $Global:AuditRunspace = $runspace

    $Global:AuditWatcher = New-Object System.Windows.Threading.DispatcherTimer
    $Global:AuditWatcher.Interval = [TimeSpan]::FromMilliseconds(400)
    $Global:AuditWatcher.Add_Tick({
        if (-not $Global:AuditHandle.IsCompleted) { return }
        $Global:AuditWatcher.Stop()

        try { [void]$Global:AuditPS.EndInvoke($Global:AuditHandle) }
        catch { Write-Log "Audit runspace ended with error: $($_.Exception.Message)" }

        try { $Global:AuditRunspace.Close(); $Global:AuditRunspace.Dispose() } catch { }
        try { $Global:AuditPS.Dispose() } catch { }

        $SearchButton.IsEnabled = $true
        $ReloadButton.IsEnabled = $true
        Update-ResultCount
    })
    $Global:AuditWatcher.Start()
}

###############################################################################
# SEARCH BUTTON
###############################################################################

$SearchButton.Add_Click({
    $selected = Get-SelectedGroup
    if (-not $selected) {
        [System.Windows.MessageBox]::Show("Please select a valid group.", "$AppName - Group Required", "OK", "Warning") | Out-Null
        return
    }

    try {
        Ensure-GraphConnection
    } catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, "$AppName - Sign-In Failed", "OK", "Error") | Out-Null
        return
    }

    Clear-Results
    $SearchButton.IsEnabled = $false
    $ReloadButton.IsEnabled = $false
    $ProgressBar.Value = 0
    Set-Status -Text "Starting audit for '$($selected.displayName)'..."

    Start-AuditRunspace -GroupId $selected.id -GroupName $selected.displayName
})

###############################################################################
# RUN
###############################################################################

$Window.ShowDialog() | Out-Null
