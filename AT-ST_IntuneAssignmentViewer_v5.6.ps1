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
    Dark = @{
        Background       = "#1F1F1F"
        Surface          = "#292929"
        Subtle           = "#333333"
        Border           = "#3D3D3D"
        TextPrimary      = "#FFFFFF"
        TextSecondary    = "#ADADAD"
        Accent           = "#115EA3"
        AccentHover      = "#0F6CBD"
        AccentPressed    = "#0C3B5E"
        DataGridAltRow   = "#2E2E2E"
        ButtonHover      = "#3D3D3D"
        ButtonPressed    = "#1F1F1F"
        HeaderBackground = "#242424"
    }
    Light = @{
        Background       = "#F5F5F5"
        Surface          = "#FFFFFF"
        Subtle           = "#F0F0F0"
        Border           = "#E0E0E0"
        TextPrimary      = "#242424"
        TextSecondary    = "#616161"
        Accent           = "#0F6CBD"
        AccentHover      = "#115EA3"
        AccentPressed    = "#0F548C"
        DataGridAltRow   = "#FAFAFA"
        ButtonHover      = "#F5F5F5"
        ButtonPressed    = "#E0E0E0"
        HeaderBackground = "#FAFAFA"
    }
}

function Get-SavedTheme {
    if (Test-Path $ThemePath) {
        try {
            $data = Get-Content $ThemePath -Raw | ConvertFrom-Json
            if ($data.Theme -in @("Dark","Light")) { return $data.Theme }
        } catch {
            Write-Log "Failed to read theme preference: $($_.Exception.Message)"
        }
    }
    return "Dark"
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
        Title="Assignment and Targeted Scope Tracker"
        Height="820"
        Width="1240"
        WindowStartupLocation="CenterScreen"
        Background="{DynamicResource BackgroundBrush}"
        FontFamily="Segoe UI Variable, Segoe UI"
        FontSize="14">

    <Window.Resources>
        <SolidColorBrush x:Key="BackgroundBrush" Color="#1F1F1F"/>
        <SolidColorBrush x:Key="SurfaceBrush" Color="#292929"/>
        <SolidColorBrush x:Key="SubtleBrush" Color="#333333"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#3D3D3D"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#ADADAD"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#115EA3"/>
        <SolidColorBrush x:Key="AccentHoverBrush" Color="#0F6CBD"/>
        <SolidColorBrush x:Key="AccentPressedBrush" Color="#0C3B5E"/>
        <SolidColorBrush x:Key="DataGridAltRowBrush" Color="#2E2E2E"/>
        <SolidColorBrush x:Key="ButtonHoverBrush" Color="#3D3D3D"/>
        <SolidColorBrush x:Key="ButtonPressedBrush" Color="#1F1F1F"/>
        <SolidColorBrush x:Key="HeaderBackgroundBrush" Color="#242424"/>

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
                    Content="Switch to Light"
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
# >>> PASTE YOUR LOGO BASE64 BETWEEN THE @' AND '@ MARKERS BELOW <<<
###############################################################################

$LogoBase64 = @'

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

    if (-not $Themes.ContainsKey($ThemeName)) { $ThemeName = "Dark" }
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

    if ($ThemeName -eq "Dark") {
        $ThemeToggleButton.Content = "Switch to Light"
    } else {
        $ThemeToggleButton.Content = "Switch to Dark"
    }

    $Global:CurrentTheme = $ThemeName
    Save-Theme -ThemeName $ThemeName
    Write-Log "Applied theme: $ThemeName"
}

Apply-Theme -ThemeName $Global:CurrentTheme

$ThemeToggleButton.Add_Click({
    if ($Global:CurrentTheme -eq "Dark") {
        Apply-Theme -ThemeName "Light"
    } else {
        Apply-Theme -ThemeName "Dark"
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
