<#
karanik_WinMaintenance.ps1 (Production)

WPF GUI launcher (download -> cache -> run) with MenuBar (File / Tools / Info).
- Default base: %ProgramData%\karanik_WinMaintenance (Logs/Cache/Temp inside)
- Settings dialog to override folders
- Download timeout + Execution timeout (prevents "freeze")
- Supports right-click "Run with PowerShell" (self relaunch in STA)
- Downloads remote scripts to cache and removes Mark-of-the-Web (Unblock-File) automatically for cached scripts
- Full logging (no silent errors, no "(no message)")

FIX: Replaced BackgroundWorker with [PowerShell]::Create() + dedicated Runspace
     to avoid "There is no Runspace available to run scripts in this thread" error.

Remote base path used for filename-only ScriptName values:
  https://karanik.gr/scripts/powershell/karanik_WinMaintenance

ScriptName formats supported:
- "script_name.ps1" (uses base above)
- "/scripts/powershell/karanik_WinMaintenance/script_name.ps1" (absolute site path)
- "scripts/powershell/karanik_WinMaintenance/script_name.ps1" (relative site path)
- "https://..." (absolute URL)

Env vars passed to child scripts:
  KARANIK_WM_LOGDIR, KARANIK_WM_CACHEDIR, KARANIK_WM_TEMPDIR, KARANIK_WM_VERBOSE
#>

#region Admin Auto-Elevation
# If not running as Administrator, relaunch elevated (UAC prompt).
# This ensures all child scripts inherit admin rights.
try {
  $currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
  if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $self = $PSCommandPath
    if (-not $self) { $self = $MyInvocation.MyCommand.Path }
    if ($self) {
      $ps = $(if (($__gc = Get-Command powershell.exe -ErrorAction SilentlyContinue)) { $__gc.Source })
      if (-not $ps) { $ps = $(if (($__gc = Get-Command powershell -ErrorAction SilentlyContinue)) { $__gc.Source }) }
      if (-not $ps) { $ps = $(if (($__gc = Get-Command pwsh -ErrorAction SilentlyContinue)) { $__gc.Source }) }
      if ($ps) {
        Start-Process -FilePath $ps -ArgumentList @("-NoProfile","-ExecutionPolicy","Bypass","-File",$self) -Verb RunAs
        exit
      }
    }
  }
} catch { }
#endregion Admin Auto-Elevation

#region STA Guard (WPF requires STA)
try {
  $apt = [System.Threading.Thread]::CurrentThread.ApartmentState
  if ($apt -ne "STA") {
    $self = $PSCommandPath
    if (-not $self) { $self = $MyInvocation.MyCommand.Path }
    if (-not $self) { throw "Cannot determine script path for STA relaunch." }

    $ps = $(if (($__gc = Get-Command powershell.exe -ErrorAction SilentlyContinue)) { $__gc.Source })
    if (-not $ps) { $ps = $(if (($__gc = Get-Command powershell -ErrorAction SilentlyContinue)) { $__gc.Source }) }
    if (-not $ps) { $ps = $(if (($__gc = Get-Command pwsh -ErrorAction SilentlyContinue)) { $__gc.Source }) }
    if (-not $ps) { throw "powershell.exe/pwsh not found." }

    $args = @("-NoProfile","-STA","-ExecutionPolicy","Bypass","-File",$self)
    Start-Process -FilePath $ps -ArgumentList $args -WindowStyle Normal | Out-Null
    exit
  }
} catch { }
#endregion STA Guard

# Best-effort: keep this launcher unblocked if possible (does not bypass policy if blocked).
try { if ($PSCommandPath) { Unblock-File -Path $PSCommandPath -ErrorAction SilentlyContinue } } catch { }

#region Global config / defaults
$ErrorActionPreference = "Stop"

$AppName  = "karanik_WinMaintenance"
$AboutUrl = "https://karanik.gr"

# Remote base (filename-only ScriptName values)
$BaseUrl = "https://karanik.gr/scripts/powershell/karanik_WinMaintenance"

# Timeouts
$DownloadTimeoutSec = 25
$ExecTimeoutSec     = 3600

$DefaultBaseDir = Join-Path $env:ProgramData $AppName
$ConfigPath     = Join-Path $DefaultBaseDir "config.json"

$State = [ordered]@{
  VerboseOutput = $false
  Theme = "Auto"
  Paths = [ordered]@{
    Base  = $DefaultBaseDir
    Logs  = (Join-Path $DefaultBaseDir "Logs")
    Cache = (Join-Path $DefaultBaseDir "Cache")
    Temp  = (Join-Path $DefaultBaseDir "Temp")
  }
}
#endregion Global config / defaults

#region Logging
function Ensure-Dirs {
  param([Parameter(Mandatory)][hashtable]$Paths)
  foreach ($p in @($Paths.Base,$Paths.Logs,$Paths.Cache,$Paths.Temp)) {
    if (-not $p) { continue }
    if (-not (Test-Path $p)) { New-Item -Path $p -ItemType Directory -Force | Out-Null }
  }
}

function NowStamp { (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") }

function Write-LogLine {
  param(
    [AllowNull()][string]$Message,
    [ValidateSet("INFO","WARN","ERROR","DEBUG","SUCCESS")][string]$Level = "INFO"
  )
  if ([string]::IsNullOrWhiteSpace($Message)) { $Message = "(no message)" }
  Ensure-Dirs -Paths $State.Paths
  $logFile = Join-Path $State.Paths.Logs ("{0}_{1:yyyyMMdd}.log" -f $AppName,(Get-Date))
  $line = "[{0}][{1}] {2}" -f (NowStamp),$Level,$Message
  try { Add-Content -Path $logFile -Value $line -Encoding UTF8 } catch { }
  return $line
}

function Format-ErrorFull {
  param([Parameter(Mandatory)]$Err)
  try {
    if ($Err -is [System.Management.Automation.ErrorRecord]) {
      $ex = $Err.Exception
      $msg = $Err | Out-String
      if (-not $msg.Trim()) { $msg = $ex.ToString() }
      return $msg.Trim()
    }
    return (($Err | Out-String).Trim())
  } catch {
    return "Unknown error (failed to format)."
  }
}
#endregion Logging

#region Network/Download/Exec
function Enable-Tls12 {
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072 } catch { }
}

function Get-FileSha256 {
  param([Parameter(Mandatory)][string]$Path)
  (Get-FileHash -Path $Path -Algorithm SHA256).Hash.ToUpperInvariant()
}

function Get-CacheFileName {
  param([Parameter(Mandatory)][string]$Url)
  # Extract the original filename from the URL (e.g. "UpdateWindows.ps1")
  # Falls back to sanitized full URL only if filename cannot be determined.
  $uriPath = $Url.Trim()
  $fileName = $null
  try {
    $uri = [System.Uri]::new($uriPath)
    $seg = $uri.Segments | Select-Object -Last 1
    $seg = [System.Uri]::UnescapeDataString($seg).Trim("/\")
    if ($seg -and $seg -match "\.ps1$") { $fileName = $seg }
  } catch { }
  if (-not $fileName) {
    # Fallback: last path component after /
    $parts = $uriPath -split "[/\\]"
    $last = ($parts | Where-Object { $_ }) | Select-Object -Last 1
    if ($last -and $last -match "\.ps1$") { $fileName = $last }
  }
  if (-not $fileName) {
    # Last resort: full sanitize (old behavior)
    $fileName = ($uriPath -replace "https?://","") -replace "[^a-zA-Z0-9\.\-_]+","_"
    if (-not $fileName.EndsWith(".ps1")) { $fileName += ".ps1" }
  }
  $fileName
}

function Get-ChildPowerShell {
  $ps = $(if (($__gc = Get-Command powershell.exe -ErrorAction SilentlyContinue)) { $__gc.Source })
  if (-not $ps) { $ps = $(if (($__gc = Get-Command powershell -ErrorAction SilentlyContinue)) { $__gc.Source }) }
  if (-not $ps) { $ps = $(if (($__gc = Get-Command pwsh -ErrorAction SilentlyContinue)) { $__gc.Source }) }
  if (-not $ps) { throw "No PowerShell executable found (powershell/pwsh)." }
  return $ps
}

function Resolve-RemoteUrl {
  param([Parameter(Mandatory)][string]$ScriptName)

  $sn = [string]$ScriptName
  $sn = $sn.Trim()
  if (-not $sn) { return $null }

  if ($sn -match '^(?i)https?://') { return $sn }
  if ($sn.StartsWith("/")) { return ("https://karanik.gr{0}" -f $sn) }
  if ($sn -match '[\\/]' ) {
    $sn2 = ($sn -replace '\\','/').TrimStart("/")
    return ("https://karanik.gr/{0}" -f $sn2)
  }
  return ("{0}/{1}" -f $BaseUrl.TrimEnd("/"), $sn)
}

function Download-RemoteScript {
  param(
    [Parameter(Mandatory)][string]$Url,
    [string]$ExpectedSha256 = $null,
    [switch]$Force
  )
  Ensure-Dirs -Paths $State.Paths
  Enable-Tls12

  $cacheName  = Get-CacheFileName -Url $Url
  $cachedPath = Join-Path $State.Paths.Cache $cacheName

  if ((-not $Force) -and (Test-Path $cachedPath)) {
    if ($ExpectedSha256) {
      $actual = Get-FileSha256 -Path $cachedPath
      if ($actual -eq $ExpectedSha256.ToUpperInvariant()) {
        try { Unblock-File -Path $cachedPath -ErrorAction SilentlyContinue } catch { }
        return $cachedPath
      }
      Remove-Item -Path $cachedPath -Force -ErrorAction SilentlyContinue
    } else {
      try { Unblock-File -Path $cachedPath -ErrorAction SilentlyContinue } catch { }
      return $cachedPath
    }
  }

  $tmp = Join-Path $State.Paths.Temp ($cacheName + "." + [guid]::NewGuid().ToString("N") + ".tmp")
  try {
    Invoke-WebRequest -Uri $Url -OutFile $tmp -UseBasicParsing -Headers @{ "Cache-Control"="no-cache" } -TimeoutSec $DownloadTimeoutSec -ErrorAction Stop | Out-Null
  } catch {
    if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    $em = (Format-ErrorFull $_)
    throw ("Download failed for {0}: {1}" -f $Url, $em)
  }

  if (-not (Test-Path $tmp)) { throw ("Download failed for {0}: file not created." -f $Url) }

  if ($ExpectedSha256) {
    $actual = Get-FileSha256 -Path $tmp
    if ($actual -ne $ExpectedSha256.ToUpperInvariant()) {
      Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
      throw "Hash verification failed. Expected $ExpectedSha256 but got $actual"
    }
  }

  Move-Item -Path $tmp -Destination $cachedPath -Force
  try { Unblock-File -Path $cachedPath -ErrorAction SilentlyContinue } catch { }
  return $cachedPath
}

function Invoke-DownloadedScript {
  param([Parameter(Mandatory)][string]$ScriptPath)

  if (-not (Test-Path $ScriptPath)) { throw "Script not found: $ScriptPath" }

  $psExe = Get-ChildPowerShell

  $env:KARANIK_WM_LOGDIR   = $State.Paths.Logs
  $env:KARANIK_WM_CACHEDIR = $State.Paths.Cache
  $env:KARANIK_WM_TEMPDIR  = $State.Paths.Temp
  $env:KARANIK_WM_VERBOSE  = ($(if ($State.VerboseOutput) { "1" } else { "0" }))

  $args = @("-NoProfile","-ExecutionPolicy","Bypass","-File",$ScriptPath)
  if ($State.VerboseOutput) { $args += "-Verbose" }

  $p = Start-Process -FilePath $psExe -ArgumentList $args -PassThru
  if (-not $p) { throw "Failed to start process: $psExe" }

  $exited = $p.WaitForExit($ExecTimeoutSec * 1000)
  if (-not $exited) {
    try { $p.Kill() } catch { }
    throw "Execution timeout after $ExecTimeoutSec sec: $ScriptPath"
  }

  return [int]$p.ExitCode
}
#endregion Network/Download/Exec

#region Config
function Load-Config {
  try {
    if (Test-Path $ConfigPath) {
      $json = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
      if ($json.Paths) {
        foreach ($k in @("Base","Logs","Cache","Temp")) {
          if ($json.Paths.$k -and ($json.Paths.$k -is [string]) -and $json.Paths.$k.Trim()) {
            $State.Paths[$k] = $json.Paths.$k
          }
        }
      }
      if ($null -ne $json.VerboseOutput) { $State.VerboseOutput = [bool]$json.VerboseOutput }
      if ($json.Theme -and $json.Theme -in @("Auto","Light","Dark")) { $State.Theme = $json.Theme }
    }
  } catch { }
}

function Save-Config {
  Ensure-Dirs -Paths $State.Paths
  $obj = [pscustomobject]@{
    VerboseOutput = $State.VerboseOutput
    Theme = $State.Theme
    Paths = [pscustomobject]@{
      Base  = $State.Paths.Base
      Logs  = $State.Paths.Logs
      Cache = $State.Paths.Cache
      Temp  = $State.Paths.Temp
    }
  }
  $obj | ConvertTo-Json -Depth 5 | Set-Content -Path $ConfigPath -Encoding UTF8
}

function Reset-ToDefaults {
  $State.Paths.Base  = $DefaultBaseDir
  $State.Paths.Logs  = (Join-Path $DefaultBaseDir "Logs")
  $State.Paths.Cache = (Join-Path $DefaultBaseDir "Cache")
  $State.Paths.Temp  = (Join-Path $DefaultBaseDir "Temp")
}

function Open-LatestLog {
  Ensure-Dirs -Paths $State.Paths
  $file = Get-ChildItem -Path $State.Paths.Logs -Filter "$AppName*.log" -ErrorAction SilentlyContinue |
          Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if ($file) { Start-Process -FilePath $file.FullName | Out-Null }
}

function Open-FolderInExplorer {
  param([string]$Path)
  if (-not $Path) { return }
  if (-not (Test-Path $Path)) {
    New-Item -Path $Path -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
  }
  try { Start-Process explorer.exe -ArgumentList $Path -ErrorAction SilentlyContinue } catch { }
}
#endregion Config

#region Catalog
#
# ID scheme:  GroupNumber.ItemNumber
#   1 = Main
#   2 = PowerShell Tools
#   3 = Cleaning Tools
#   4 = Export Tools
#   5 = Fix Tools
#   6 = Extra Tools
#
$Catalog = @(
  # ── 1. Main ────────────────────────────────────────────────────────────────
  [pscustomobject]@{ Id="1.1";  Title="Update Windows";                                 Type="Remote"; ScriptName="UpdateWindows.ps1";   Sha256=$null; Group="Main" },
  [pscustomobject]@{ Id="1.2";  Title="Reset Windows Update";                           Type="Remote"; ScriptName="ResetWindowsUpdate.ps1";     Sha256=$null; Group="Main" },
  [pscustomobject]@{ Id="1.3";  Title="Defender: Update signatures & Quick Scan";       Type="Remote"; ScriptName="DefenderUpdate.ps1";     Sha256=$null; Group="Main" },
  [pscustomobject]@{ Id="1.4";  Title="GPUpdate /force (Computer + User)";              Type="Remote"; ScriptName="GPUpdate.ps1";     Sha256=$null; Group="Main" },
  [pscustomobject]@{ Id="1.5";  Title="Create System Restore Point";                    Type="Remote"; ScriptName="CreateRestorePoint.ps1"; Sha256=$null; Group="Main" },
  [pscustomobject]@{ Id="1.51"; Title="Clean All Restore Points";                        Type="Remote"; ScriptName="CleanRestorePoints.ps1";  Sha256=$null; Group="Main" },
  [pscustomobject]@{ Id="1.6";  Title="Clear Event Logs";                               Type="Remote"; ScriptName="ClearEventLogs.ps1";     Sha256=$null; Group="Main" },
  [pscustomobject]@{ Id="1.7";  Title="Install Winget (App Installer helper)";          Type="Remote"; ScriptName="InstallWinget.ps1";     Sha256=$null; Group="Main" },
  [pscustomobject]@{ Id="1.8";  Title="Install latest PowerShell";                      Type="Remote"; ScriptName="InstallPowerShell.ps1";     Sha256=$null; Group="Main" },
  [pscustomobject]@{ Id="1.9";  Title="Upgrade all apps via Winget (excluding pinned)"; Type="Remote"; ScriptName="UpgradeWingetApps.ps1";     Sha256=$null; Group="Main" },

  # ── 2. PowerShell Tools ────────────────────────────────────────────────────
  [pscustomobject]@{ Id="2.1"; Title="PS Module Manager"; 							Type="Remote"; ScriptName="https://karanik.gr/scripts/powershell/PSModuleManager/PSModuleManager.ps1"; Sha256=$null; Group="PowerShell Tools" },

  # ── 3. Cleaning Tools ──────────────────────────────────────────────────────
  [pscustomobject]@{ Id="3.1"; Title="Clean all Temporary files (system & users)"; Type="Remote"; ScriptName="CleanTempFiles.ps1";                  Sha256=$null; Group="Cleaning Tools" },
  [pscustomobject]@{ Id="3.2"; Title="Clear Teams cache only";                     Type="Remote"; ScriptName="ClearTeamsCache.ps1";                  Sha256=$null; Group="Cleaning Tools" },
  [pscustomobject]@{ Id="3.3"; Title="Teams Full clean + reinstall";   Type="Standalone"; ScriptName="TeamsCleanupAndInstall.ps1";       Sha256=$null; Group="Cleaning Tools" },
  [pscustomobject]@{ Id="3.4"; Title="Cleanup SoftwareDistribution";               Type="Remote"; ScriptName="Cleanup-SoftwareDistribution.ps1";    Sha256=$null; Group="Cleaning Tools" },

  # ── 4. Export Tools ────────────────────────────────────────────────────────
  [pscustomobject]@{ Id="4.1"; Title="System Information (show + save full report)";       Type="Remote"; ScriptName="SystemInfo.ps1";                  Sha256=$null; Group="Export Tools" },
  [pscustomobject]@{ Id="4.2"; Title="Export Installed Software list (CSV to Desktop)";    Type="Remote"; ScriptName="ExportInstalledSoftware.ps1";      Sha256=$null; Group="Export Tools" },
  [pscustomobject]@{ Id="4.3"; Title="Export Drivers (CSV to Desktop)";                    Type="Remote"; ScriptName="ExportDrivers.ps1";                Sha256=$null; Group="Export Tools" },
  [pscustomobject]@{ Id="4.4"; Title="Wi-Fi Profiles: Export to Desktop (XML, key=clear)"; Type="Remote"; ScriptName="ExportWiFiProfiles.ps1";           Sha256=$null; Group="Export Tools" },
  [pscustomobject]@{ Id="4.5"; Title="Wi-Fi Profiles: Import from Desktop XML";            Type="Remote"; ScriptName="ImportWiFiProfiles.ps1";           Sha256=$null; Group="Export Tools" },
  [pscustomobject]@{ Id="4.6"; Title="Scheduled Tasks Audit (CSV to Desktop + console)";   Type="Remote"; ScriptName="ExportScheduledTasks.ps1";         Sha256=$null; Group="Export Tools" },
  [pscustomobject]@{ Id="4.7"; Title="Windows Update History Summary (CSV + console)";     Type="Remote"; ScriptName="ExportWindowsUpdateHistory.ps1";   Sha256=$null; Group="Export Tools" },
  [pscustomobject]@{ Id="4.8"; Title="Battery Report";                                     Type="Remote"; ScriptName="BatteryReport.ps1";                Sha256=$null; Group="Export Tools" },

  # ── 5. Fix Tools ───────────────────────────────────────────────────────────
  [pscustomobject]@{ Id="5.1"; Title="Reset Print Spooler";                          Type="Remote"; ScriptName="ResetPrintSpooler.ps1";         Sha256=$null; Group="Fix Tools" },
  [pscustomobject]@{ Id="5.2"; Title="Network Reset";                                Type="Remote"; ScriptName="NetworkReset.ps1";               Sha256=$null; Group="Fix Tools" },
  [pscustomobject]@{ Id="5.3"; Title="Wi-Fi Reset";                                  Type="Remote"; ScriptName="WiFiReset.ps1";                  Sha256=$null; Group="Fix Tools" },
  [pscustomobject]@{ Id="5.4"; Title="Set Classic Right-Click Menu (Win11)";         Type="Remote"; ScriptName="SetClassicContextMenu.ps1";     Sha256=$null; Group="Fix Tools" },
  [pscustomobject]@{ Id="5.5"; Title="Restore Default Right-Click Menu (Win11)";     Type="Remote"; ScriptName="RestoreDefaultContextMenu.ps1"; Sha256=$null; Group="Fix Tools" },
  [pscustomobject]@{ Id="5.6"; Title="Disable Hibernation";                          Type="Remote"; ScriptName="DisableHibernation.ps1";             Sha256=$null; Group="Fix Tools" },
  [pscustomobject]@{ Id="5.7"; Title="Enable Hibernation";                           Type="Remote"; ScriptName="EnableHibernation.ps1";              Sha256=$null; Group="Fix Tools" },
  [pscustomobject]@{ Id="5.8"; Title="WebView2_Repair_Install Phase1"; 				 Type="Remote";  ScriptName="WebView2_Repair_Install.ps1";    Sha256=$null; Group="Fix Tools" },  
  [pscustomobject]@{ Id="5.9"; Title="WebView2_Repair_Install Phase2"; 				 Type="Remote";  ScriptName="WebView2_Repair_Install_v3.ps1";    Sha256=$null; Group="Fix Tools" },    
  [pscustomobject]@{ Id="5.10"; Title="Remove_iCloud_AllUsers"; 					 Type="Remote";  ScriptName="Remove_iCloud_AllUsers.ps1";    Sha256=$null; Group="Fix Tools" },    

  # ── 6. Extra Tools ─────────────────────────────────────────────────────────
  [pscustomobject]@{ Id="6.1"; Title="Reset RDS Licensing Grace Period (Windows Server)"; Type="Remote";  ScriptName="RDS-GracePeriod-Manager.ps1";    Sha256=$null; Group="Extra Tools" },
  [pscustomobject]@{ Id="6.2"; Title="Microsoft Activation Scripts (MAS)";                Type="Inline";  Command="irm https://get.activated.win | iex"; Sha256=$null; Group="Extra Tools" },
  [pscustomobject]@{ Id="6.3"; Title="WinScript";                                         Type="Inline";  Command='irm "https://winscript.cc/irm" | iex'; Sha256=$null; Group="Extra Tools" },
  [pscustomobject]@{ Id="6.4"; Title="Chris Titus Tech's Windows Utility";                Type="Inline";  Command='irm "https://christitus.com/win" | iex'; Sha256=$null; Group="Extra Tools"; Tooltip=$null },
  [pscustomobject]@{ Id="6.5"; Title="WinDiag-AI"; 										  Type="Remote";  ScriptName="https://karanik.gr/scripts/powershell/WinDiag-AI/WinDiag-AI.ps1";    Sha256=$null; Group="Extra Tools" },
  
  # -- 7. Office Tools --
  [pscustomobject]@{ Id="7.1"; Title="Office Tools"; Type="Standalone"; ScriptName="Manage-OfficeUpdates.ps1"; Sha256=$null; Group="Office Tools";
    Tooltip="Microsoft 365 / Office Click-to-Run Manager + SaRA Enterprise`n`nUPDATE CONTROL`n  1. Enable automatic updates   (registry policy = 1)`n  2. Disable automatic updates  (registry policy = 0)`n  3. Check update status and version`n  4. Run update now (OfficeC2RClient /runnow)`n`nREPAIR`n  5a. Quick Repair  - local, no internet required`n  5b. Online Repair - full re-download from Microsoft`n`nCONTROL PANEL`n  6.  Open Mail (Microsoft Outlook) - mlcfg32.cpl`n`nSaRA ENTERPRISE (Microsoft Support and Recovery Assistant)`n  7.  Uninstall ALL versions of Office`n  8.  Uninstall specific version (M365 / 2021 / 2019 / 2016 ...)`n  9.  Outlook Scan - diagnostics and config report`n  10. Reset Office Activation - clears licenses and cached accounts`n  11. Fix Office Activation issues - automated recovery`n  12. Fix Teams Meeting Add-in for Outlook`n  13. Outlook Calendar Scan (CalCheck)" }
)


function Strip-MenuPrefix {
  param([string]$Title)
  if (-not $Title) { return $Title }
  return ($Title -replace '^\s*\[\d+(\.\d+)*\]\s*', '')
}

function Get-AllRunnableCatalogItems {
  $Catalog | Where-Object { $_.Type -eq "Remote" }
}
#endregion Catalog

#region WPF UI
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Load-Config
Ensure-Dirs -Paths $State.Paths

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="karanik_WinMaintenance"
        Height="740" Width="1160"
        MinHeight="480" MinWidth="700"
        WindowStartupLocation="CenterScreen">
  <Window.Resources>

    <Style x:Key="ToolbarButtonStyle" TargetType="Button">
      <Setter Property="Height" Value="28"/>
      <Setter Property="Padding" Value="12,0"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderBrush" Value="#CCCCCC"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="5" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#F0F0F0"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#E0E0E0"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="RunButtonStyle" TargetType="Button" BasedOn="{StaticResource ToolbarButtonStyle}">
      <Setter Property="Background" Value="#1E6EB5"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderBrush" Value="#1558A0"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="5" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#1558A0"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#0F4580"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="ToggleButtonStyle" TargetType="ToggleButton">
      <Setter Property="Height" Value="28"/>
      <Setter Property="Padding" Value="12,0"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderBrush" Value="#CCCCCC"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ToggleButton">
            <Border x:Name="bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="5" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#E8F0FA"/>
                <Setter TargetName="bd" Property="BorderBrush" Value="#1E6EB5"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#F0F0F0"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="FilterBoxStyle" TargetType="TextBox">
      <Setter Property="Height" Value="26"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Padding" Value="8,0"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="BorderBrush" Value="#CCCCCC"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border Background="White" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="5">
              <ScrollViewer x:Name="PART_ContentHost" Margin="2,0"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="TreeLeafStyle" TargetType="TreeViewItem">
      <Setter Property="Padding" Value="4,4,4,4"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TreeViewItem">
            <StackPanel>
              <Border x:Name="Bd" Background="{TemplateBinding Background}"
                      BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="0" CornerRadius="4"
                      Padding="{TemplateBinding Padding}"
                      Margin="2,1,2,1">
                <ContentPresenter x:Name="PART_Header" ContentSource="Header"
                                  HorizontalAlignment="Left"/>
              </Border>
              <ItemsPresenter x:Name="ItemsHost"/>
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#DCEAF8"/>
                <Setter Property="Foreground" Value="#1558A0"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="#F2F2F2"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="TreeGroupStyle" TargetType="TreeViewItem">
      <Setter Property="Padding" Value="6,6,6,6"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="Foreground" Value="#555555"/>
      <Setter Property="IsExpanded" Value="True"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TreeViewItem">
            <StackPanel>
              <Border x:Name="GrpBd" Margin="0,8,0,0"
                      Background="#F2F4F7"
                      BorderBrush="#D8DCE3"
                      BorderThickness="0,1,0,1"
                      Padding="{TemplateBinding Padding}">
                <DockPanel>
                  <Path x:Name="Arrow" DockPanel.Dock="Left"
                        Width="10" Height="10" Margin="0,0,7,0"
                        VerticalAlignment="Center" StrokeThickness="2"
                        StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                        Stroke="#666666" Fill="Transparent"
                        Data="M 1 3 L 5 7 L 9 3"/>
                  <ContentPresenter ContentSource="Header" VerticalAlignment="Center"/>
                </DockPanel>
              </Border>
              <ItemsPresenter x:Name="ItemsHost"/>
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsExpanded" Value="False">
                <Setter TargetName="Arrow" Property="Data" Value="M 3 1 L 7 5 L 3 9"/>
                <Setter TargetName="ItemsHost" Property="Visibility" Value="Collapsed"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="GrpBd" Property="Background" Value="#E8EBF0"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Menu Grid.Row="0" Background="#F8F8F8" BorderBrush="#E0E0E0" BorderThickness="0,0,0,1">
      <MenuItem Header="_File">
        <MenuItem x:Name="mnuSettings" Header="Settings..."/>
        <MenuItem x:Name="mnuOpenLog" Header="Open log..."/>
        <Separator/>
        <MenuItem x:Name="mnuReload" Header="Reload"/>
        <Separator/>
        <MenuItem x:Name="mnuExit" Header="Exit"/>
      </MenuItem>
      <MenuItem Header="_Tools">
        <MenuItem x:Name="mnuRun" Header="Run selected"/>
        <MenuItem x:Name="mnuRunSerial" Header="Run Serial..."/>
        <Separator/>
        <MenuItem x:Name="mnuForce" Header="Force re-download" IsCheckable="True"/>
        <MenuItem x:Name="mnuVerbose" Header="Verbose output" IsCheckable="True"/>
      </MenuItem>
      <MenuItem Header="_Info">
        <MenuItem x:Name="mnuAbout" Header="About"/>
      </MenuItem>
    </Menu>

    <Border Grid.Row="1" Background="#F8F8F8" BorderBrush="#E8E8E8" BorderThickness="0,0,0,1" Padding="12,8">
      <DockPanel>
        <StackPanel DockPanel.Dock="Left" Orientation="Vertical" VerticalAlignment="Center">
          <TextBlock FontSize="16" FontWeight="SemiBold" Text="karanik_WinMaintenance" Foreground="#1A1A1A"/>
          <TextBlock FontSize="11" Foreground="#888888" Text="Online launcher  --  download -&gt; cache -&gt; run" Margin="0,1,0,0"/>
        </StackPanel>

        <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center" HorizontalAlignment="Right">
          <Button x:Name="btnClean" Style="{StaticResource ToolbarButtonStyle}"
                  Width="32" Height="28" Padding="0" ToolTip="Clean folders"
                  Content="&#x1F5D1;" FontSize="14"/>
          <Popup x:Name="popClean" PlacementTarget="{Binding ElementName=btnClean}"
                 Placement="Bottom" StaysOpen="False" AllowsTransparency="True">
            <Border Background="White" BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="6"
                    Padding="4" Effect="{x:Null}">
              <StackPanel Width="140">
                <Button x:Name="btnCleanLogs"  Content="&#x1F5D1;  Clean Logs"  Height="28" Margin="0,1" Style="{StaticResource ToolbarButtonStyle}" HorizontalContentAlignment="Left" Padding="10,0"/>
                <Button x:Name="btnCleanCache" Content="&#x1F5D1;  Clean Cache" Height="28" Margin="0,1" Style="{StaticResource ToolbarButtonStyle}" HorizontalContentAlignment="Left" Padding="10,0"/>
                <Button x:Name="btnCleanTemp"  Content="&#x1F5D1;  Clean Temp"  Height="28" Margin="0,1" Style="{StaticResource ToolbarButtonStyle}" HorizontalContentAlignment="Left" Padding="10,0"/>
              </StackPanel>
            </Border>
          </Popup>
          <Separator Width="1" Background="#D0D0D0" Margin="6,4,6,4"/>
          <Button x:Name="btnTheme" Style="{StaticResource ToolbarButtonStyle}"
                  Width="32" Height="28" Padding="0" ToolTip="Theme"
                  Content="&#x2600;" FontSize="14"/>
          <Popup x:Name="popTheme" PlacementTarget="{Binding ElementName=btnTheme}"
                 Placement="Bottom" StaysOpen="False" AllowsTransparency="True">
            <Border Background="White" BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="6"
                    Padding="4" Effect="{x:Null}">
              <StackPanel Width="130">
                <Button x:Name="btnThemeAuto"  Content="Auto (system)"  Height="28" Margin="0,1" Style="{StaticResource ToolbarButtonStyle}" HorizontalContentAlignment="Left" Padding="10,0"/>
                <Button x:Name="btnThemeLight" Content="Light"          Height="28" Margin="0,1" Style="{StaticResource ToolbarButtonStyle}" HorizontalContentAlignment="Left" Padding="10,0"/>
                <Button x:Name="btnThemeDark"  Content="Dark"           Height="28" Margin="0,1" Style="{StaticResource ToolbarButtonStyle}" HorizontalContentAlignment="Left" Padding="10,0"/>
              </StackPanel>
            </Border>
          </Popup>
        </StackPanel>

        <Rectangle/>
      </DockPanel>
    </Border>

    <Border Grid.Row="2" Background="#FAFAFA" BorderBrush="#E8E8E8" BorderThickness="0,0,0,1" Padding="12,6">
      <DockPanel>
        <Button x:Name="btnRun" DockPanel.Dock="Left"
                Style="{StaticResource RunButtonStyle}"
                Content="Run  &#x25B6;" Padding="14,0" Margin="0,0,8,0"/>
        <Button x:Name="btnRunSerial" DockPanel.Dock="Left"
                Style="{StaticResource ToolbarButtonStyle}"
                Content="Run Serial..." Margin="0,0,8,0">
          <Button.ToolTip>
            <ToolTip MaxWidth="300">
              <TextBlock TextWrapping="Wrap">
                <Run FontWeight="Bold">Run Serial</Run>
                <LineBreak/>
                Runs all selected scripts one after another, in order.
                Each script waits to finish before the next one starts.
                Useful for running a full maintenance sequence unattended.
              </TextBlock>
            </ToolTip>
          </Button.ToolTip>
        </Button>
        <ToggleButton x:Name="tglForce" DockPanel.Dock="Left"
                      Style="{StaticResource ToggleButtonStyle}"
                      Content="Force re-download" Margin="0,0,6,0">
          <ToggleButton.ToolTip>
            <ToolTip MaxWidth="300">
              <TextBlock TextWrapping="Wrap">
                <Run FontWeight="Bold">Force re-download</Run>
                <LineBreak/>
                When enabled, scripts are always downloaded fresh from the server
                even if a cached version already exists locally.
                Use this after uploading an updated version of a script.
              </TextBlock>
            </ToolTip>
          </ToggleButton.ToolTip>
        </ToggleButton>
        <ToggleButton x:Name="tglVerbose" DockPanel.Dock="Left"
                      Style="{StaticResource ToggleButtonStyle}"
                      Content="Verbose" Margin="0,0,0,0">
          <ToggleButton.ToolTip>
            <ToolTip MaxWidth="300">
              <TextBlock TextWrapping="Wrap">
                <Run FontWeight="Bold">Verbose</Run>
                <LineBreak/>
                When enabled, scripts output additional debug information
                including detailed step-by-step progress and diagnostic messages.
                Useful for troubleshooting script execution.
              </TextBlock>
            </ToolTip>
          </ToggleButton.ToolTip>
        </ToggleButton>
        <ToggleButton x:Name="tglStealth" DockPanel.Dock="Left"
                      Style="{StaticResource ToggleButtonStyle}"
                      Content="&#x1F575; Stealth Mode" Margin="8,0,0,0"
                      Foreground="#8B0000" FontWeight="SemiBold">
          <ToggleButton.ToolTip>
            <ToolTip MaxWidth="320">
              <TextBlock TextWrapping="Wrap">
                <Run FontWeight="Bold">Stealth Mode</Run>
                <LineBreak/>
                When enabled and you close the application, ALL files and folders
                created by karanik_WinMaintenance will be permanently deleted
                (Logs, Cache, Temp and the entire C:\ProgramData\karanik_WinMaintenance folder).
                <LineBreak/>
                <Run Foreground="#8B0000" FontStyle="Italic">Files are deleted permanently  -  not sent to Recycle Bin.</Run>
              </TextBlock>
            </ToolTip>
          </ToggleButton.ToolTip>
        </ToggleButton>

        <Button x:Name="btnClearLog" DockPanel.Dock="Right"
                Style="{StaticResource ToolbarButtonStyle}"
                Content="Clear log" Margin="0,0,0,0"/>

        <TextBlock x:Name="txtStatus" DockPanel.Dock="Right"
                   VerticalAlignment="Center" FontSize="12"
                   Foreground="#555555" Margin="0,0,12,0"
                   HorizontalAlignment="Right"/>
      </DockPanel>
    </Border>

    <Grid Grid.Row="3" Margin="12,10,12,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="300" MinWidth="160"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*" MinWidth="200"/>
      </Grid.ColumnDefinitions>

      <Border Grid.Column="0" BorderBrush="#DCDCDC" BorderThickness="1" CornerRadius="8">
        <DockPanel>
          <Border DockPanel.Dock="Top" Padding="10,8,10,6" BorderBrush="#ECECEC" BorderThickness="0,0,0,1">
            <TextBox x:Name="txtFilter" Style="{StaticResource FilterBoxStyle}"
                     ToolTip="Filter actions..."/>
          </Border>
          <TreeView x:Name="tvMenu" BorderThickness="0" Background="Transparent"
                    Padding="4,4,4,4" ScrollViewer.HorizontalScrollBarVisibility="Disabled"/>
        </DockPanel>
      </Border>

      <GridSplitter Grid.Column="1" Width="7" HorizontalAlignment="Center"
                    VerticalAlignment="Stretch" Background="Transparent"
                    Margin="2,0,2,0" ShowsPreview="False"
                    ResizeBehavior="PreviousAndNext" ResizeDirection="Columns"
                    Cursor="SizeWE"/>

      <Grid Grid.Column="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="2*" MinHeight="60"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*" MinHeight="60"/>
        </Grid.RowDefinitions>

        <!-- Log output (top pane) -->
        <Border Grid.Row="0" BorderBrush="#DCDCDC" BorderThickness="1" CornerRadius="8">
          <RichTextBox x:Name="txtOutput"
                       FontFamily="Consolas" FontSize="11.5"
                       IsReadOnly="True"
                       VerticalScrollBarVisibility="Auto"
                       HorizontalScrollBarVisibility="Disabled"
                       Padding="10,8" BorderThickness="0" Background="Transparent"
                       IsDocumentEnabled="True">
            <FlowDocument/>
          </RichTextBox>
        </Border>

        <!-- Resizable splitter between log and terminal -->
        <GridSplitter Grid.Row="1"
                      Height="7"
                      HorizontalAlignment="Stretch"
                      VerticalAlignment="Center"
                      Background="Transparent"
                      Margin="0,2,0,2"
                      ShowsPreview="False"
                      ResizeBehavior="PreviousAndNext"
                      ResizeDirection="Rows"
                      Cursor="SizeNS"/>

        <!-- Terminal output (bottom pane) -->
        <Border Grid.Row="2" BorderBrush="#DCDCDC" BorderThickness="1" CornerRadius="8">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Border Grid.Row="0" Background="#F8F8F8" BorderBrush="#ECECEC" BorderThickness="0,0,0,1"
                    Padding="10,4,10,4" CornerRadius="8,8,0,0">
              <DockPanel>
                <TextBlock Text="Terminal Output" FontSize="10" FontWeight="SemiBold"
                           Foreground="#888888" VerticalAlignment="Center"/>
                <Button x:Name="btnClearTerminal" DockPanel.Dock="Right"
                        Content="Clear" FontSize="10" Height="20" Padding="8,0"
                        Background="Transparent" BorderBrush="#CCCCCC" BorderThickness="1"
                        Cursor="Hand" HorizontalAlignment="Right"/>
              </DockPanel>
            </Border>
            <RichTextBox Grid.Row="1" x:Name="txtTerminal"
                     FontFamily="Consolas" FontSize="11.5"
                     IsReadOnly="True"
                     VerticalScrollBarVisibility="Auto"
                     HorizontalScrollBarVisibility="Disabled"
                     Padding="10,8" BorderThickness="0" Background="Transparent"
                     IsDocumentEnabled="True">
              <FlowDocument/>
            </RichTextBox>
          </Grid>
        </Border>
      </Grid>
    </Grid>

    <TextBlock Grid.Row="4" Margin="12,6,12,8" Foreground="#AAAAAA" FontSize="11"
               Text="Default base: %ProgramData%\karanik_WinMaintenance  |  You can override paths in Settings."/>
  </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$Window = [Windows.Markup.XamlReader]::Load($reader)

$tvMenu          = $Window.FindName("tvMenu")
$txtOutput       = $Window.FindName("txtOutput")
$txtTerminal     = $Window.FindName("txtTerminal")
$txtStatus       = $Window.FindName("txtStatus")
$txtFilter       = $Window.FindName("txtFilter")
$btnClearTerminal = $Window.FindName("btnClearTerminal")

$mnuSettings   = $Window.FindName("mnuSettings")
$mnuOpenLog    = $Window.FindName("mnuOpenLog")
$mnuReload     = $Window.FindName("mnuReload")
$mnuExit       = $Window.FindName("mnuExit")
$mnuRun        = $Window.FindName("mnuRun")
$mnuRunSerial  = $Window.FindName("mnuRunSerial")
$mnuForce      = $Window.FindName("mnuForce")
$mnuVerbose    = $Window.FindName("mnuVerbose")
$mnuAbout      = $Window.FindName("mnuAbout")

$btnRun        = $Window.FindName("btnRun")
$btnRunSerial  = $Window.FindName("btnRunSerial")
$tglForce      = $Window.FindName("tglForce")
$tglVerbose    = $Window.FindName("tglVerbose")
$tglStealth    = $Window.FindName("tglStealth")
$btnClearLog   = $Window.FindName("btnClearLog")
$btnClean      = $Window.FindName("btnClean")
$btnCleanLogs  = $Window.FindName("btnCleanLogs")
$btnCleanCache = $Window.FindName("btnCleanCache")
$btnCleanTemp  = $Window.FindName("btnCleanTemp")
$popClean      = $Window.FindName("popClean")
$btnTheme      = $Window.FindName("btnTheme")
$popTheme      = $Window.FindName("popTheme")
$btnThemeAuto  = $Window.FindName("btnThemeAuto")
$btnThemeLight = $Window.FindName("btnThemeLight")
$btnThemeDark  = $Window.FindName("btnThemeDark")

function Ui-Append {
  param([string]$Line)
  $color = $null
  if     ($Line -match '\[ERROR\]'   -or $Line -match '\[ERR\]')     { $color = [System.Windows.Media.Brushes]::Red }
  elseif ($Line -match '\[WARN\]'    -or $Line -match '\[WARNING\]') { $color = [System.Windows.Media.Brushes]::DarkOrange }
  elseif ($Line -match '\[SUCCESS\]' -or $Line -match '\[OK\]')     { $color = [System.Windows.Media.Brushes]::Green }
  elseif ($Line -match '\[DEBUG\]')  { $color = [System.Windows.Media.Brushes]::Gray }
  elseif ($Line -match '\[SCAN\]'    -or $Line -match '\[RESULT\]') { $color = [System.Windows.Media.Brushes]::DodgerBlue }
  elseif ($Line -match '\[BATCH\]')  { $color = [System.Windows.Media.Brushes]::MediumOrchid }

  $para = $txtOutput.Document.Blocks | Select-Object -Last 1
  if (-not $para -or $para -isnot [System.Windows.Documents.Paragraph]) {
    $para = [System.Windows.Documents.Paragraph]::new()
    $para.Margin = [System.Windows.Thickness]::new(0)
    $txtOutput.Document.Blocks.Add($para)
  }
  $run = [System.Windows.Documents.Run]::new($Line + [Environment]::NewLine)
  if ($color) { $run.Foreground = $color }
  $para.Inlines.Add($run)
  $txtOutput.ScrollToEnd()
}

function Ui-ClearLog { $txtOutput.Document.Blocks.Clear() }

function Ui-AppendTerminal {
  param([string]$Line)
  if (-not $txtTerminal) { return }
  $color = $null
  if     ($Line -match '\[ERROR\]'   -or $Line -match '\[ERR\]'  -or $Line -match 'ERROR ')  { $color = [System.Windows.Media.Brushes]::Red }
  elseif ($Line -match '\[WARN\]'    -or $Line -match '\[WARNING\]') { $color = [System.Windows.Media.Brushes]::DarkOrange }
  elseif ($Line -match '\[SUCCESS\]' -or $Line -match '\[OK\]')     { $color = [System.Windows.Media.Brushes]::Green }
  elseif ($Line -match '\[DEBUG\]')  { $color = [System.Windows.Media.Brushes]::Gray }
  elseif ($Line -match '\[SCAN\]'    -or $Line -match '\[RESULT\]') { $color = [System.Windows.Media.Brushes]::DodgerBlue }
  elseif ($Line -match '\[BATCH\]')  { $color = [System.Windows.Media.Brushes]::MediumOrchid }
  elseif ($Line -match '\[INFO\]')   { $color = [System.Windows.Media.Brushes]::CornflowerBlue }

  $para = $txtTerminal.Document.Blocks | Select-Object -Last 1
  if (-not $para -or $para -isnot [System.Windows.Documents.Paragraph]) {
    $para = [System.Windows.Documents.Paragraph]::new()
    $para.Margin = [System.Windows.Thickness]::new(0)
    $txtTerminal.Document.Blocks.Add($para)
  }
  $run = [System.Windows.Documents.Run]::new($Line + [Environment]::NewLine)
  if ($color) { $run.Foreground = $color }
  $para.Inlines.Add($run)
  $txtTerminal.ScrollToEnd()
}

function Ui-ClearTerminal {
  if ($txtTerminal) { $txtTerminal.Document.Blocks.Clear() }
}

function Ui-SetStatus {
  param([string]$s)
  $txtStatus.Text = $s
  # Color the status text: green for success, red for error, grey otherwise
  if ($s -match "success|ExitCode: 0") {
    $txtStatus.Foreground = [System.Windows.Media.Brushes]::DarkGreen
  } elseif ($s -match "error|Error|ExitCode: [^0]") {
    $txtStatus.Foreground = [System.Windows.Media.Brushes]::DarkRed
  } else {
    $txtStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x55,0x55,0x55))
  }
}

function Ui-Log { param([string]$msg,[string]$level="INFO") Ui-Append (Write-LogLine -Message $msg -Level $level) }

function Set-UiBusy([bool]$Busy) {
  # Settings / Open log / Exit / Tree stay ALWAYS enabled
  $mnuReload.IsEnabled     = -not $Busy
  $mnuRun.IsEnabled        = -not $Busy
  $mnuRunSerial.IsEnabled  = -not $Busy
  $mnuForce.IsEnabled      = -not $Busy
  $mnuVerbose.IsEnabled    = -not $Busy
  $btnRun.IsEnabled        = -not $Busy
  $btnRunSerial.IsEnabled  = -not $Busy
  $tglForce.IsEnabled      = -not $Busy
  $tglVerbose.IsEnabled    = -not $Busy
  $txtFilter.IsEnabled     = -not $Busy
  if ($btnClearTerminal) { $btnClearTerminal.IsEnabled = -not $Busy }
}

# Tracks which catalog items have been cached (URL -> cached path)
$Script:CachedItems = [System.Collections.Generic.Dictionary[string,string]]::new()

function Build-Tree {
  param([string]$Filter = "")
  $tvMenu.Items.Clear()
  # Clear multi-selection state when tree is rebuilt
  if ($Script:MultiSelected) { $Script:MultiSelected.Clear() }

  $order = @("Main","PowerShell Tools","Cleaning Tools","Export Tools","Fix Tools","Extra Tools","Office Tools")
  foreach ($grpName in $order) {
    $items = @($Catalog | Where-Object Group -eq $grpName | Sort-Object { [double]($_.Id -replace '[^0-9.]','') })

    # Apply filter (force-expand all when filtering)
    $forceExpand = $false
    if ($Filter.Trim()) {
      $items = @($items | Where-Object { $_.Title -match [regex]::Escape($Filter.Trim()) })
      $forceExpand = $true
    }
    if ($items.Count -lt 1) { continue }

    $node            = New-Object System.Windows.Controls.TreeViewItem
    $node.Header     = $grpName.ToUpper()
    $node.Style      = $Window.FindResource("TreeGroupStyle")
    $node.IsExpanded = $true
    $node.Tag        = $grpName

    # Collapse/expand when clicking the group header.
    # We check that no leaf TreeViewItem is an ancestor of the click source  - 
    # that way clicks on child items do NOT trigger collapse.
    $node.Add_PreviewMouseLeftButtonDown({
      param($src,$e)
      # Walk up from OriginalSource  -  if we find a leaf TVI before the group node, abort
      $el = $e.OriginalSource
      while ($el -and $el -ne $src) {
        if ($el -is [System.Windows.Controls.TreeViewItem] -and $el.Tag) {
          return  # click was inside a leaf - do nothing
        }
        $el = [System.Windows.Media.VisualTreeHelper]::GetParent($el)
      }
      # Click was on the group header area  -  toggle expand
      $src.IsExpanded = -not $src.IsExpanded
      $e.Handled = $true
    })

    foreach ($item in $items) {
      $leaf        = New-Object System.Windows.Controls.TreeViewItem
      $leaf.Header = $item.Title
      $leaf.Tag    = $item
      $leaf.Style  = $Window.FindResource("TreeLeafStyle")

      # Tooltip from catalog entry
      if ($item.Tooltip) {
        $tt            = New-Object System.Windows.Controls.ToolTip
        $tt.MaxWidth   = 420
        $tt.Content    = $item.Tooltip
        $tt.FontFamily = "Consolas"
        $tt.FontSize   = 11
        $leaf.ToolTip  = $tt
      }

      # Double-click runs the item immediately
      $leaf.Add_MouseDoubleClick({
        param($src,$e)
        $clickedItem = $src.Tag
        if (-not $clickedItem) { return }
        $e.Handled = $true
        try {
          if ($clickedItem.Type -eq "Picker") { Show-FixToolsPicker; return }
          Start-RunRemoteItem -Item $clickedItem
        } catch {
          Ui-Append (Write-LogLine -Message (Format-ErrorFull $_) -Level "ERROR")
          Ui-SetStatus "Error."
          Set-UiBusy $false
        }
      })

      # Right-click context menu
      $ctxMenu = New-Object System.Windows.Controls.ContextMenu

      $ctxRun = New-Object System.Windows.Controls.MenuItem
      $ctxRun.Header = "Run"
      $ctxRun.FontWeight = "SemiBold"
      $ctxRun.Tag = $item
      $ctxRun.Add_Click({
        param($s,$e)
        $ci = $s.Tag
        if (-not $ci) { return }
        try {
          if ($ci.Type -eq "Picker") { Show-FixToolsPicker; return }
          Start-RunRemoteItem -Item $ci
        } catch {
          Ui-Append (Write-LogLine -Message (Format-ErrorFull $_) -Level "ERROR")
          Ui-SetStatus "Error."
          Set-UiBusy $false
        }
      })
      $ctxMenu.Items.Add($ctxRun) | Out-Null

      $ctxAddSerial = New-Object System.Windows.Controls.MenuItem
      $ctxAddSerial.Header = "Add to Run Serial..."
      $ctxAddSerial.Tag = $item
      $ctxAddSerial.Add_Click({
        param($s,$e)
        $ci = $s.Tag
        if (-not $ci) { return }

        # Collect the right-clicked item + any Ctrl+click multi-selected items
        $toAdd = [System.Collections.Generic.List[pscustomobject]]::new()
        $toAdd.Add($ci)

        # Pull from the MultiSelected dictionary (populated by Ctrl+click)
        if ($Script:MultiSelected -and $Script:MultiSelected.Count -gt 0) {
          foreach ($kvp in @($Script:MultiSelected.GetEnumerator())) {
            if ($kvp.Key -ne $ci.Id -and $kvp.Value.Tag) {
              $toAdd.Add($kvp.Value.Tag)
            }
          }
        }

        # Deduplicate
        $unique = [System.Collections.Generic.List[pscustomobject]]::new()
        $seen = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($it in $toAdd) {
          if ($seen.Add($it.Id)) { $unique.Add($it) }
        }

        Ui-Log ("Opening Run Serial with {0} pre-selected item(s)..." -f $unique.Count) "INFO"
        try {
          $items = Show-SerialRunDialog -PreSelected $unique.ToArray()
          if ($items) { Start-RunBatch -Items $items }
        } catch {
          Ui-Append (Write-LogLine -Message (Format-ErrorFull $_) -Level "ERROR")
          Ui-SetStatus "Error."
          Set-UiBusy $false
        }
      })
      $ctxMenu.Items.Add($ctxAddSerial) | Out-Null

      $ctxSep = New-Object System.Windows.Controls.Separator
      $ctxMenu.Items.Add($ctxSep) | Out-Null

      $ctxShowScript = New-Object System.Windows.Controls.MenuItem
      $ctxShowScript.Header = "Show cached script"
      $ctxShowScript.Tag = $item
      $ctxShowScript.Add_Click({
        param($s,$e)
        $ci = $s.Tag
        if (-not $ci -or -not $ci.ScriptName) { return }

        $resolvedUrl = if ($ci.Type -ne "Inline") { Resolve-RemoteUrl -ScriptName $ci.ScriptName } else { "" }
        if (-not $resolvedUrl) { return }

        # Build expected cache filename using same logic as Get-CacheFileName
        $cachedPath = Join-Path $State.Paths.Cache (Get-CacheFileName -Url $resolvedUrl)

        if (Test-Path $cachedPath) {
          Ui-Log ("Opening cached script: {0}" -f $cachedPath) "INFO"
          try { Start-Process -FilePath "explorer.exe" -ArgumentList ("/select,`"{0}`"" -f $cachedPath) }
          catch { Start-Process -FilePath $State.Paths.Cache }
        } else {
          [System.Windows.MessageBox]::Show(
            ("Script has not been downloaded yet.`n`nIt will be cached here when first run:`n{0}" -f $cachedPath),
            "Not cached yet", "OK", "Information") | Out-Null
        }
      })
      $ctxMenu.Items.Add($ctxShowScript) | Out-Null

      $leaf.ContextMenu = $ctxMenu
      [void]$node.Items.Add($leaf)
    }
    [void]$tvMenu.Items.Add($node)
  }

  $tvMenu.UpdateLayout()
  # Re-apply current theme so new items get correct colors (guard: function defined later)
  if (Get-Command Apply-Theme -ErrorAction SilentlyContinue) { Apply-Theme $State.Theme }
}

# Multi-select implementation using a manual selection set.
# Ctrl+click toggles a leaf into/out of $Script:MultiSelected (Dictionary: Id -> TVI)
# and highlights it with a blue tint. Normal click (no Ctrl) clears the set.
$Script:MultiSelected = [System.Collections.Generic.Dictionary[string,object]]::new()

$tvMenu.Add_PreviewMouseLeftButtonDown({
  param($src, $e)

  # Find the leaf TVI under cursor
  $hit = $src.InputHitTest($e.GetPosition($src))
  $tvi = $hit
  while ($tvi -and $tvi -isnot [System.Windows.Controls.TreeViewItem]) {
    $tvi = [System.Windows.Media.VisualTreeHelper]::GetParent($tvi)
  }

  $isLeaf = ($tvi -and $tvi.Tag -and $tvi.Tag -isnot [string])
  $ctrlDown = ([System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::LeftCtrl) -or
               [System.Windows.Input.Keyboard]::IsKeyDown([System.Windows.Input.Key]::RightCtrl))

  if (-not $ctrlDown) {
    # Normal click  -  clear all manual selections
    foreach ($kvp in @($Script:MultiSelected.GetEnumerator())) {
      try { $kvp.Value.Background = [System.Windows.Media.Brushes]::Transparent } catch {}
    }
    $Script:MultiSelected.Clear()
    # Let normal WPF selection happen
    return
  }

  # Ctrl is down  -  only act on leaf nodes
  if (-not $isLeaf) { return }

  $id = $tvi.Tag.Id
  if ($Script:MultiSelected.ContainsKey($id)) {
    # Deselect this item
    $Script:MultiSelected.Remove($id) | Out-Null
    $tvi.IsSelected = $false
    $tvi.Background = [System.Windows.Media.Brushes]::Transparent
  } else {
    # Select this item
    $Script:MultiSelected[$id] = $tvi
    $tvi.IsSelected = $true
    $tvi.Background = [System.Windows.Media.SolidColorBrush]::new(
      [System.Windows.Media.Color]::FromRgb(0xDC,0xEA,0xF8))
  }
  $e.Handled = $true   # prevent TreeView from deselecting everything else
})

function Get-SelectedLeafItem {
  $sel = $tvMenu.SelectedItem
  if (-not $sel) { return $null }
  if ($sel -is [System.Windows.Controls.TreeViewItem] -and $sel.Tag) { return $sel.Tag }
  return $null
}

Build-Tree
$mnuVerbose.IsChecked  = [bool]$State.VerboseOutput
$mnuForce.IsChecked    = $false
$tglVerbose.IsChecked  = [bool]$State.VerboseOutput
$tglForce.IsChecked    = $false

# Apply saved theme on startup
if ($State.Theme -ne "Auto") {
  $Window.Dispatcher.InvokeAsync({
    Apply-Theme $State.Theme
  }) | Out-Null
}

Ui-SetStatus "Ready."
Ui-Log "Started $AppName. BaseUrl=$BaseUrl"
Ui-Log ("Timeouts: Download={0}s | Exec={1}s" -f $DownloadTimeoutSec,$ExecTimeoutSec)
Ui-Log ("Paths: Base={0} | Logs={1} | Cache={2} | Temp={3}" -f $State.Paths.Base,$State.Paths.Logs,$State.Paths.Cache,$State.Paths.Temp)
#endregion WPF UI

#region ─── WORKER (FIX: [PowerShell]::Create() + Runspace instead of BackgroundWorker) ───
#
# ROOT CAUSE OF ORIGINAL BUG:
#   BackgroundWorker.DoWork runs on a .NET ThreadPool thread that has NO PowerShell
#   Runspace attached. Any PS cmdlet called from that thread (Add-Content, Get-Date,
#   Invoke-WebRequest, etc.) throws:
#     "There is no Runspace available to run scripts in this thread."
#
# FIX APPROACH:
#   1. Build a dedicated Runspace and inject all required variables + helper functions.
#   2. Run the job inside [PowerShell]::Create() bound to that Runspace via BeginInvoke().
#   3. Use a WPF DispatcherTimer (runs on UI thread) to poll completion and drain a
#      thread-safe ConcurrentQueue<string> used for progress messages.
#   This keeps the UI fully responsive and gives the worker a proper PS environment.
#
$Script:_jobPs     = $null   # [PowerShell] instance for current job
$Script:_jobRs     = $null   # [Runspace]   instance for current job
$Script:_jobHandle = $null   # IAsyncResult from BeginInvoke
$Script:_jobQueue  = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$Script:_jobTimer  = $null   # DispatcherTimer

# ScriptBlock containing all helper functions that must exist inside the worker Runspace.
# Keep in sync with the main-thread versions above.
$Script:WorkerHelpers = {
  # ---- injected via Runspace variables: $State, $DownloadTimeoutSec, $ExecTimeoutSec,
  #      $AppName, $BaseUrl, $ProgressQueue ----

  function NowStamp { [datetime]::Now.ToString("yyyy-MM-dd HH:mm:ss") }

  function Ensure-Dirs {
    param([hashtable]$Paths)
    foreach ($p in @($Paths.Base,$Paths.Logs,$Paths.Cache,$Paths.Temp)) {
      if (-not $p) { continue }
      if (-not [System.IO.Directory]::Exists($p)) {
        [System.IO.Directory]::CreateDirectory($p) | Out-Null
      }
    }
  }

  function Write-LogLine {
    param([string]$Message,[string]$Level="INFO")
    if ([string]::IsNullOrWhiteSpace($Message)) { $Message = "(no message)" }
    Ensure-Dirs -Paths $State.Paths
    $logFile = [System.IO.Path]::Combine($State.Paths.Logs,
                 ("{0}_{1:yyyyMMdd}.log" -f $AppName,[datetime]::Now))
    $line = "[{0}][{1}] {2}" -f (NowStamp),$Level,$Message
    try { [System.IO.File]::AppendAllText($logFile, $line + "`r`n", [System.Text.Encoding]::UTF8) } catch { }
    return $line
  }

  function Report {
    param([string]$msg,[string]$lvl="INFO")
    $line = Write-LogLine -Message $msg -Level $lvl
    $ProgressQueue.Enqueue($line)
  }

  function Format-ErrorFull {
    param($Err)
    try {
      if ($Err -is [System.Management.Automation.ErrorRecord]) {
        $msg = $Err | Out-String
        if (-not $msg.Trim()) { $msg = $Err.Exception.ToString() }
        return $msg.Trim()
      }
      return (($Err | Out-String).Trim())
    } catch { return "Unknown error (failed to format)." }
  }

  function Enable-Tls12 {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072 } catch { }
  }

  function Get-FileSha256 {
    param([string]$Path)
    (Get-FileHash -Path $Path -Algorithm SHA256).Hash.ToUpperInvariant()
  }

  function Get-CacheFileName {
    param([string]$Url)
    $uriPath = $Url.Trim()
    $fileName = $null
    try {
      $uri = [System.Uri]::new($uriPath)
      $seg = $uri.Segments | Select-Object -Last 1
      $seg = [System.Uri]::UnescapeDataString($seg).Trim("/\")
      if ($seg -and $seg -match "\.ps1$") { $fileName = $seg }
    } catch { }
    if (-not $fileName) {
      $parts = $uriPath -split "[/\\]"
      $last = ($parts | Where-Object { $_ }) | Select-Object -Last 1
      if ($last -and $last -match "\.ps1$") { $fileName = $last }
    }
    if (-not $fileName) {
      $fileName = ($uriPath -replace "https?://","") -replace "[^a-zA-Z0-9\.\-_]+","_"
      if (-not $fileName.EndsWith(".ps1")) { $fileName += ".ps1" }
    }
    $fileName
  }

  function Get-ChildPowerShell {
    $ps = $(if (($__gc = Get-Command powershell.exe -ErrorAction SilentlyContinue)) { $__gc.Source })
    if (-not $ps) { $ps = $(if (($__gc = Get-Command powershell -ErrorAction SilentlyContinue)) { $__gc.Source }) }
    if (-not $ps) { $ps = $(if (($__gc = Get-Command pwsh -ErrorAction SilentlyContinue)) { $__gc.Source }) }
    if (-not $ps) { throw "No PowerShell executable found (powershell/pwsh)." }
    return $ps
  }

  function Download-RemoteScript {
    param([string]$Url,[string]$ExpectedSha256=$null,[bool]$Force=$false)
    Ensure-Dirs -Paths $State.Paths
    Enable-Tls12

    $cacheName  = Get-CacheFileName -Url $Url
    $cachedPath = [System.IO.Path]::Combine($State.Paths.Cache, $cacheName)

    if ((-not $Force) -and [System.IO.File]::Exists($cachedPath)) {
      if ($ExpectedSha256) {
        $actual = Get-FileSha256 -Path $cachedPath
        if ($actual -eq $ExpectedSha256.ToUpperInvariant()) {
          try { Unblock-File -Path $cachedPath -ErrorAction SilentlyContinue } catch { }
          return $cachedPath
        }
        Remove-Item -Path $cachedPath -Force -ErrorAction SilentlyContinue
      } else {
        try { Unblock-File -Path $cachedPath -ErrorAction SilentlyContinue } catch { }
        return $cachedPath
      }
    }

    $tmp = [System.IO.Path]::Combine($State.Paths.Temp,
             ($cacheName + "." + [guid]::NewGuid().ToString("N") + ".tmp"))
    try {
      Invoke-WebRequest -Uri $Url -OutFile $tmp -UseBasicParsing `
        -Headers @{ "Cache-Control"="no-cache" } `
        -TimeoutSec $DownloadTimeoutSec -ErrorAction Stop | Out-Null
    } catch {
      if ([System.IO.File]::Exists($tmp)) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
      throw ("Download failed for {0}: {1}" -f $Url,(Format-ErrorFull $_))
    }

    if (-not [System.IO.File]::Exists($tmp)) { throw ("Download failed for {0}: file not created." -f $Url) }

    if ($ExpectedSha256) {
      $actual = Get-FileSha256 -Path $tmp
      if ($actual -ne $ExpectedSha256.ToUpperInvariant()) {
        Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
        throw "Hash verification failed. Expected $ExpectedSha256 but got $actual"
      }
    }

    Move-Item -Path $tmp -Destination $cachedPath -Force
    try { Unblock-File -Path $cachedPath -ErrorAction SilentlyContinue } catch { }
    return $cachedPath
  }

  function Invoke-DownloadedScript {
    param(
      [string]$ScriptPath,
      [string]$LogBaseName = ""
    )
    if (-not [System.IO.File]::Exists($ScriptPath)) { throw "Script not found: $ScriptPath" }

    $psExe = Get-ChildPowerShell

    $cleanBase    = if ($LogBaseName) { $LogBaseName } else { [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath) }
    $childLogName = "{0}_{1}.log" -f $cleanBase, [datetime]::Now.ToString("yyyyMMdd_HHmmss")
    $childLogPath = [System.IO.Path]::Combine($State.Paths.Logs, $childLogName)

    $pargs = [System.Collections.Generic.List[string]]::new()
    $pargs.Add("-NoProfile")
    $pargs.Add("-ExecutionPolicy")
    $pargs.Add("Bypass")
    $pargs.Add("-File")
    $pargs.Add($ScriptPath)
    if ($State.VerboseOutput) { $pargs.Add("-Verbose") }

    # Redirect stdout/stderr to a temp text file instead of pipes.
    # This avoids the "no Runspace" crash with event handlers and
    # the blocking EndOfStream issue with synchronous ReadLine.
    $stdioFile = [System.IO.Path]::Combine($State.Paths.Temp,
                   $cleanBase + "_stdio_" + [datetime]::Now.ToString("yyyyMMdd_HHmmss") + ".txt")

    $psi = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName               = $psExe
    $psi.Arguments              = ($pargs | ForEach-Object { if ($_ -match '\s') { "`"$_`"" } else { $_ } }) -join " "
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $false
    # NO stdout/stderr redirect — avoids all Runspace/blocking issues
    $psi.RedirectStandardOutput = $false
    $psi.RedirectStandardError  = $false

    # Pass env vars directly into the child process environment
    $psi.EnvironmentVariables["KARANIK_WM_LOGDIR"]   = $State.Paths.Logs
    $psi.EnvironmentVariables["KARANIK_WM_CACHEDIR"]  = $State.Paths.Cache
    $psi.EnvironmentVariables["KARANIK_WM_TEMPDIR"]   = $State.Paths.Temp
    $psi.EnvironmentVariables["KARANIK_WM_VERBOSE"]   = $(if ($State.VerboseOutput) { "1" } else { "0" })
    $psi.EnvironmentVariables["KARANIK_WM_LOGFILE"]   = $childLogPath

    $ProgressQueue.Enqueue("Tailing log: $childLogPath")

    $proc = [System.Diagnostics.Process]::Start($psi)
    if (-not $proc) { throw "Failed to start process: $psExe" }

    # Pure log-file-tailing loop. No stdout/stderr involvement.
    $tailPos  = 0L
    $pollMs   = 300
    $started  = [datetime]::Now

    while ($true) {
      $elapsed = ([datetime]::Now - $started).TotalMilliseconds
      if ($elapsed -ge ($ExecTimeoutSec * 1000)) {
        try { $proc.Kill() } catch { }
        throw "Execution timeout after $ExecTimeoutSec sec: $ScriptPath"
      }

      # Tail child log file -> ProgressQueue (shows in txtOutput via Ui-Append)
      if ([System.IO.File]::Exists($childLogPath)) {
        try {
          $fs      = [System.IO.File]::Open($childLogPath, [System.IO.FileMode]::Open,
                                            [System.IO.FileAccess]::Read,
                                            [System.IO.FileShare]::ReadWrite)
          $fs.Seek($tailPos, [System.IO.SeekOrigin]::Begin) | Out-Null
          $sr      = [System.IO.StreamReader]::new($fs, [System.Text.Encoding]::UTF8)
          $newText = $sr.ReadToEnd()
          $tailPos = $fs.Position
          $sr.Dispose()
          $fs.Dispose()
          if ($newText.Length -gt 0) {
            foreach ($ln in ($newText -split "`r?`n")) {
              if ($ln.Trim()) { $ProgressQueue.Enqueue($ln) }
            }
          }
        } catch { }
      }

      if ($proc.HasExited) { break }
      [System.Threading.Thread]::Sleep($pollMs)
    }

    # Final log tail drain after exit
    if ([System.IO.File]::Exists($childLogPath)) {
      try {
        $fs      = [System.IO.File]::Open($childLogPath, [System.IO.FileMode]::Open,
                                          [System.IO.FileAccess]::Read,
                                          [System.IO.FileShare]::ReadWrite)
        $fs.Seek($tailPos, [System.IO.SeekOrigin]::Begin) | Out-Null
        $sr      = [System.IO.StreamReader]::new($fs, [System.Text.Encoding]::UTF8)
        $newText = $sr.ReadToEnd()
        $sr.Dispose()
        $fs.Dispose()
        if ($newText.Length -gt 0) {
          foreach ($ln in ($newText -split "`r?`n")) {
            if ($ln.Trim()) { $ProgressQueue.Enqueue($ln) }
          }
        }
      } catch { }
    }

    return [int]$proc.ExitCode
  }
}

# ScriptBlock that performs the actual work (runs inside the dedicated Runspace).
$Script:WorkerJob = {
  param($job)

  try {
    if ($job.Type -eq "Remote") {
      Report ("Selected: {0}" -f $job.Title) "INFO"
      Report ("Resolved URL: {0}" -f $job.Url) "DEBUG"
      Report "Downloading..." "INFO"
      $local = Download-RemoteScript -Url $job.Url -ExpectedSha256 $job.Sha256 -Force $job.Force
      Report ("Cached: {0}" -f $local) "INFO"
      Report "Executing..." "INFO"
      # Extract clean script name from URL for the log filename (e.g. "UpdateWindows" from ".../UpdateWindows.ps1")
      $cleanName = [System.IO.Path]::GetFileNameWithoutExtension(($job.Url -split "[/?]" | Select-Object -Last 1))
      $exit = Invoke-DownloadedScript -ScriptPath $local -LogBaseName $cleanName
      Report ("ExitCode: {0}" -f $exit) "INFO"
      return $exit
    }

    if ($job.Type -eq "Batch") {
      $anyFail = $false
      foreach ($sub in $job.Items) {
        Report "----" "INFO"
        Report ("Batch item: {0}" -f $sub.Title) "INFO"
        try {
          Report ("Resolved URL: {0}" -f $sub.Url) "DEBUG"
          $local = Download-RemoteScript -Url $sub.Url -ExpectedSha256 $sub.Sha256 -Force $job.Force
          Report ("Cached: {0}" -f $local) "INFO"
          $cleanName2 = [System.IO.Path]::GetFileNameWithoutExtension(($sub.Url -split "[/?]" | Select-Object -Last 1))
          $exit = Invoke-DownloadedScript -ScriptPath $local -LogBaseName $cleanName2
          Report ("ExitCode: {0}" -f $exit) "INFO"
          if ($exit -ne 0) { $anyFail = $true }
        } catch {
          Report (Format-ErrorFull $_) "ERROR"
          $anyFail = $true
        }
      }
      return $(if ($anyFail) { 1 } else { 0 })
    }

    if ($job.Type -eq "Standalone") {
      # Standalone: WPF GUI script - needs UseShellExecute=$true for desktop session.
      # We can't redirect stdout with UseShellExecute=true, so we pass env vars
      # via EncodedCommand wrapper that sets them then calls the script.
      Report ("Selected: {0}" -f $job.Title) "INFO"
      Report ("Resolved URL: {0}" -f $job.Url) "DEBUG"
      Report "Downloading..." "INFO"
      $local = Download-RemoteScript -Url $job.Url -ExpectedSha256 $job.Sha256 -Force $job.Force
      Report ("Cached: {0}" -f $local) "INFO"
      Report "Launching in separate window..." "INFO"

      $psExe = Get-ChildPowerShell
      $cleanName    = [System.IO.Path]::GetFileNameWithoutExtension(($job.Url -split "[/?]" | Select-Object -Last 1))
      $childLogPath = [System.IO.Path]::Combine($State.Paths.Logs, ("{0}_{1}.log" -f $cleanName, [datetime]::Now.ToString("yyyyMMdd_HHmmss")))
      $verboseVal   = if ($State.VerboseOutput) { "1" } else { "0" }

      # Build a small PS command that sets env vars then dot-sources the script
      $innerCmd = (
          "[System.Environment]::SetEnvironmentVariable('KARANIK_WM_LOGFILE','{0}','Process');" +
          "[System.Environment]::SetEnvironmentVariable('KARANIK_WM_LOGDIR','{1}','Process');" +
          "[System.Environment]::SetEnvironmentVariable('KARANIK_WM_VERBOSE','{2}','Process');" +
          "& '{3}'"
      ) -f $childLogPath.Replace("'","''"),
            $State.Paths.Logs.Replace("'","''"),
            $verboseVal,
            $local.Replace("'","''")

      $encoded = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($innerCmd))

      $psi = [System.Diagnostics.ProcessStartInfo]::new($psExe)
      $psi.Arguments       = "-NoProfile -STA -ExecutionPolicy Bypass -EncodedCommand $encoded"
      $psi.UseShellExecute = $true
      $psi.Verb            = "runas"

      try {
        $proc = [System.Diagnostics.Process]::Start($psi)
        if ($proc) {
          Report ("Standalone window launched (PID {0})." -f $proc.Id) "INFO"
          Report ("Tailing log: {0}" -f $childLogPath) "INFO"

          # Tail the child log file and forward lines to the launcher UI.
          # UseShellExecute=true means we cannot redirect stdout/stderr,
          # but the child writes structured log lines to $childLogPath which
          # we can tail-poll while it runs.
          $tailPos  = 0L
          $pollMs   = 300
          $started  = [datetime]::Now

          while ($true) {
            $elapsed = ([datetime]::Now - $started).TotalMilliseconds
            if ($elapsed -ge ($ExecTimeoutSec * 1000)) {
              Report ("Standalone timeout after {0}s." -f $ExecTimeoutSec) "WARN"
              break
            }

            # Tail log file -> ProgressQueue (same mechanism as Remote scripts)
            if ([System.IO.File]::Exists($childLogPath)) {
              try {
                $fs      = [System.IO.File]::Open($childLogPath,
                             [System.IO.FileMode]::Open,
                             [System.IO.FileAccess]::Read,
                             [System.IO.FileShare]::ReadWrite)
                $fs.Seek($tailPos, [System.IO.SeekOrigin]::Begin) | Out-Null
                $sr      = [System.IO.StreamReader]::new($fs, [System.Text.Encoding]::UTF8)
                $newText = $sr.ReadToEnd()
                $tailPos = $fs.Position
                $sr.Dispose(); $fs.Dispose()
                if ($newText.Length -gt 0) {
                  foreach ($ln in ($newText -split "`r?`n")) {
                    if ($ln.Trim()) { $ProgressQueue.Enqueue($ln) }
                  }
                }
              } catch { }
            }

            if ($proc.HasExited) { break }
            [System.Threading.Thread]::Sleep($pollMs)
          }

          # Final drain after process exits
          if ([System.IO.File]::Exists($childLogPath)) {
            try {
              $fs      = [System.IO.File]::Open($childLogPath,
                           [System.IO.FileMode]::Open,
                           [System.IO.FileAccess]::Read,
                           [System.IO.FileShare]::ReadWrite)
              $fs.Seek($tailPos, [System.IO.SeekOrigin]::Begin) | Out-Null
              $sr      = [System.IO.StreamReader]::new($fs, [System.Text.Encoding]::UTF8)
              $newText = $sr.ReadToEnd()
              $sr.Dispose(); $fs.Dispose()
              foreach ($ln in ($newText -split "`r?`n")) {
                if ($ln.Trim()) { $ProgressQueue.Enqueue($ln) }
              }
            } catch { }
          }

          $exitCode = try { [int]$proc.ExitCode } catch { 0 }
          Report ("Standalone exited. Code: {0}" -f $exitCode) "INFO"
          return $exitCode
        }
      } catch {
        Report ("Failed to launch standalone: {0}" -f $_.Exception.Message) "ERROR"
        return 1
      }
      return 0
    }

    if ($job.Type -eq "Inline") {
      Report ("Selected: {0}" -f $job.Title) "INFO"
      Report ("Command : {0}" -f $job.Command) "DEBUG"
      # Inline commands (irm ... | iex) run in a new elevated PS window so they get
      # their own interactive session and can show their own UI.
      $psExe = $(if (($gc = Get-Command powershell.exe -ErrorAction SilentlyContinue)) { $gc.Source })
      if (-not $psExe) { $psExe = "powershell.exe" }
      $encodedCmd = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($job.Command))
      $psi = [System.Diagnostics.ProcessStartInfo]::new($psExe)
      $psi.Arguments       = "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedCmd"
      $psi.UseShellExecute = $true   # show window, allow interactive UI
      $psi.Verb            = "runas" # ensure elevated
      try {
        $proc = [System.Diagnostics.Process]::Start($psi)
        if ($proc) {
          Report ("Launched: {0}" -f $job.Title) "INFO"
          $proc.WaitForExit()
          Report ("Process exited with code: {0}" -f $proc.ExitCode) "INFO"
          return $proc.ExitCode
        }
      } catch {
        Report ("Failed to launch inline command: {0}" -f $_.Exception.Message) "ERROR"
        return 1
      }
      return 0
    }

    throw "Unknown job type: $($job.Type)"

  } catch {
    Report (Format-ErrorFull $_) "ERROR"
    return 1
  }
}

function Start-WorkerJob {
  param([pscustomobject]$Job)

  # Safety: don't start a second job if one is already running
  if ($Script:_jobTimer -and $Script:_jobTimer.IsEnabled) {
    Ui-Log "A job is already running. Please wait." "WARN"
    return
  }

  # Clear old queue
  $Script:_jobQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

  # Build a fresh Runspace with its own PS session state
  $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
  $Script:_jobRs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
  $Script:_jobRs.ApartmentState = [System.Threading.ApartmentState]::MTA
  $Script:_jobRs.ThreadOptions   = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
  $Script:_jobRs.Open()

  # Inject shared variables into the Runspace
  $Script:_jobRs.SessionStateProxy.SetVariable("State",              $State)
  $Script:_jobRs.SessionStateProxy.SetVariable("DownloadTimeoutSec", $DownloadTimeoutSec)
  $Script:_jobRs.SessionStateProxy.SetVariable("ExecTimeoutSec",     $ExecTimeoutSec)
  $Script:_jobRs.SessionStateProxy.SetVariable("AppName",            $AppName)
  $Script:_jobRs.SessionStateProxy.SetVariable("BaseUrl",            $BaseUrl)
  $Script:_jobRs.SessionStateProxy.SetVariable("ProgressQueue",      $Script:_jobQueue)

  # Load helper functions into the Runspace
  $initPs = [System.Management.Automation.PowerShell]::Create()
  $initPs.Runspace = $Script:_jobRs
  [void]$initPs.AddScript($Script:WorkerHelpers)
  try   { [void]$initPs.Invoke() }
  catch { Ui-Log ("Failed to initialise worker runspace: {0}" -f $_.Exception.Message) "ERROR"; return }
  finally { $initPs.Dispose() }

  # Start the actual job asynchronously
  $Script:_jobPs = [System.Management.Automation.PowerShell]::Create()
  $Script:_jobPs.Runspace = $Script:_jobRs
  [void]$Script:_jobPs.AddScript($Script:WorkerJob).AddArgument($Job)
  $Script:_jobHandle = $Script:_jobPs.BeginInvoke()

  # DispatcherTimer polls on the UI thread  -  no cross-thread UI access needed
  $Script:_jobTimer = New-Object System.Windows.Threading.DispatcherTimer
  $Script:_jobTimer.Interval = [System.TimeSpan]::FromMilliseconds(200)

  $Script:_jobTimer.Add_Tick({
    # Drain the progress queue -> route log lines to txtOutput, terminal lines to txtTerminal
    $line = $null
    while ($Script:_jobQueue.TryDequeue([ref]$line)) {
      if ($line.StartsWith("TERMINAL|")) {
        Ui-AppendTerminal ($line.Substring(9))
      } else {
        Ui-Append $line
      }
    }

    # Check if the PS job has finished
    if ($Script:_jobHandle -and $Script:_jobHandle.IsCompleted) {
      $Script:_jobTimer.Stop()

      # Drain any remaining messages
      $line2 = $null
      while ($Script:_jobQueue.TryDequeue([ref]$line2)) {
        if ($line2.StartsWith("TERMINAL|")) {
          Ui-AppendTerminal ($line2.Substring(9))
        } else {
          Ui-Append $line2
        }
      }

      # Collect result / errors
      $exitCode = 1
      try {
        $results = $Script:_jobPs.EndInvoke($Script:_jobHandle)
        if ($results -and $results.Count -gt 0) {
          $exitCode = [int]($results | Select-Object -Last 1)
        }
        # Surface any PS stream errors
        foreach ($err in $Script:_jobPs.Streams.Error) {
          Ui-Append (Write-LogLine -Message ($err | Out-String).Trim() -Level "ERROR")
        }
      } catch {
        Ui-Append (Write-LogLine -Message ($_.Exception.Message) -Level "ERROR")
      } finally {
        try { $Script:_jobPs.Dispose()  } catch { }
        try { $Script:_jobRs.Close()    } catch { }
        try { $Script:_jobRs.Dispose()  } catch { }
        $Script:_jobPs     = $null
        $Script:_jobRs     = $null
        $Script:_jobHandle = $null
      }

      Set-UiBusy $false
      if ($exitCode -eq 0) { Ui-SetStatus "Completed successfully." }
      else                 { Ui-SetStatus ("Completed with errors (ExitCode: {0})." -f $exitCode) }
    }
  })

  $Script:_jobTimer.Start()
}
#endregion Worker

#region Actions
function Start-RunRemoteItem($Item) {
  $force = [bool]$mnuForce.IsChecked -or [bool]$tglForce.IsChecked

  Ui-Log ("Selected: {0}" -f $Item.Title) "INFO"
  Ui-Log ("DEBUG ItemType={0} | ScriptName={1}" -f $Item.Type, ($Item.ScriptName | Out-String).Trim()) "DEBUG"

  # Inline commands (irm | iex)  -  bypass download/cache, run directly
  if ($Item.Type -eq "Standalone" -or $Item.Type -eq "Inline") {
    if ($Item.Type -eq "Standalone") {
      if (-not $Item.ScriptName) { throw ("ScriptName is null for Standalone item: {0}" -f $Item.Title) }
      $url = Resolve-RemoteUrl -ScriptName $Item.ScriptName
      if (-not $url) { throw ("Failed to resolve URL: {0}" -f $Item.ScriptName) }
      Set-UiBusy $true; Ui-SetStatus "Running..."
      $job = [pscustomobject]@{ Type="Standalone"; Title=$Item.Title; Url=$url; Sha256=$Item.Sha256; Force=$force }
      Start-WorkerJob -Job $job
      return
    }
  }
  if ($Item.Type -eq "Inline") {
    if (-not $Item.Command) { throw ("Command is null/empty for Inline item: {0}" -f $Item.Title) }
    Set-UiBusy $true
    Ui-SetStatus "Running..."
    $job = [pscustomobject]@{
      Type    = "Inline"
      Title   = $Item.Title
      Command = $Item.Command
      Force   = $false
    }
    Start-WorkerJob -Job $job
    return
  }

  if (-not $Item.ScriptName) { throw ("ScriptName is null/empty for item: {0}" -f $Item.Title) }

  $url = Resolve-RemoteUrl -ScriptName $Item.ScriptName
  if (-not $url) { throw ("Failed to resolve URL from ScriptName: {0}" -f $Item.ScriptName) }

  Set-UiBusy $true
  Ui-SetStatus "Running..."

  $job = [pscustomobject]@{
    Type  = "Remote"
    Title = $Item.Title
    Url   = $url
    Sha256= $Item.Sha256
    Force = $force
  }
  Start-WorkerJob -Job $job
}

function Start-RunBatch([array]$Items) {
  $force = [bool]$mnuForce.IsChecked -or [bool]$tglForce.IsChecked
  $batchItems = @()

  foreach ($it in $Items) {
    if ($it.Type -eq "Inline" -or $it.Type -eq "Standalone") { continue }   # cannot be batched
    if (-not $it.ScriptName) { continue }
    $u = Resolve-RemoteUrl -ScriptName $it.ScriptName
    if (-not $u) { continue }
    $batchItems += [pscustomobject]@{
      Title = (Strip-MenuPrefix $it.Title)
      Url   = $u
      Sha256= $it.Sha256
    }
  }

  if ($batchItems.Count -lt 1) { return }
  Ui-Log ("Run Serial selection count: {0}" -f $batchItems.Count) "INFO"
  Set-UiBusy $true
  Ui-SetStatus "Running..."

  $job = [pscustomobject]@{
    Type  = "Batch"
    Items = $batchItems
    Force = $force
  }
  Start-WorkerJob -Job $job
}
#endregion Actions

#region Dialogs
# ─── Dialog helpers ────────────────────────────────────────────────────────────
# Apply the main window theme (background/foreground) to a child dialog.
function New-StyledDialog {
  param(
    [string]$Title,
    [int]$Width,
    [int]$Height,
    [switch]$NoResize
  )
  $dlg = New-Object System.Windows.Window
  $dlg.Title  = $Title
  $dlg.Width  = $Width
  $dlg.Height = $Height
  $dlg.WindowStartupLocation = "CenterOwner"
  $dlg.Owner  = $Window
  $dlg.Background = [System.Windows.Media.SolidColorBrush]::new(
    [System.Windows.Media.Color]::FromRgb(0xFF,0xFF,0xFF))
  $dlg.FontFamily = [System.Windows.Media.FontFamily]::new("Segoe UI")
  $dlg.FontSize   = 12
  if ($NoResize) { $dlg.ResizeMode = "NoResize" }
  return $dlg
}

# Styled button factory  -  matches toolbar button look
function New-DlgButton {
  param(
    [string]$Label,
    [int]$Width    = 90,
    [int]$Height   = 28,
    [switch]$Primary,
    [switch]$Danger
  )

  if ($Primary) {
    $bgNormal  = "#1E6EB5"
    $bgHover   = "#1558A0"
    $bgPress   = "#0F4580"
    $fgColor   = "#FFFFFF"
    $bdColor   = "#1558A0"
    $fwValue   = "SemiBold"
  } elseif ($Danger) {
    $bgNormal  = "#C0392B"
    $bgHover   = "#A93226"
    $bgPress   = "#922B21"
    $fgColor   = "#FFFFFF"
    $bdColor   = "#A93226"
    $fwValue   = "Normal"
  } else {
    $bgNormal  = "#FFFFFF"
    $bgHover   = "#F0F0F0"
    $bgPress   = "#E0E0E0"
    $fgColor   = "#1A1A1A"
    $bdColor   = "#BBBBBB"
    $fwValue   = "Normal"
  }

  [xml]$btnXaml = @"
<Button xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="$Width" Height="$Height"
        FontSize="12" FontWeight="$fwValue"
        Foreground="$fgColor" Cursor="Hand">
  <Button.Template>
    <ControlTemplate TargetType="Button">
      <Border x:Name="bd"
              Background="$bgNormal"
              BorderBrush="$bdColor"
              BorderThickness="1"
              CornerRadius="5"
              Padding="10,0">
        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter TargetName="bd" Property="Background" Value="$bgHover"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
          <Setter TargetName="bd" Property="Background" Value="$bgPress"/>
        </Trigger>
        <Trigger Property="IsEnabled" Value="False">
          <Setter Property="Opacity" Value="0.45"/>
        </Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate>
  </Button.Template>
</Button>
"@

  $btn = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $btnXaml))
  $btn.Content = $Label
  return $btn
}

# Styled ListBox  -  clean border, no focus rectangle
function New-DlgListBox {
  param([string]$SelectionMode = "Extended")
  $lb = New-Object System.Windows.Controls.ListBox
  $lb.SelectionMode  = $SelectionMode
  $lb.BorderBrush    = [System.Windows.Media.SolidColorBrush]::new(
    [System.Windows.Media.Color]::FromRgb(0xDC,0xDC,0xDC))
  $lb.BorderThickness = "1"
  $lb.Background     = [System.Windows.Media.Brushes]::White
  $lb.FontSize       = 12
  $lb.Padding        = "2,4,2,4"
  return $lb
}

# Styled TextBox
function New-DlgTextBox {
  param([string]$Value = "", [int]$Height = 26)
  $tb = New-Object System.Windows.Controls.TextBox
  $tb.Text    = $Value
  $tb.Height  = $Height
  $tb.FontSize = 12
  $tb.VerticalContentAlignment = "Center"
  $tb.BorderBrush = [System.Windows.Media.SolidColorBrush]::new(
    [System.Windows.Media.Color]::FromRgb(0xCC,0xCC,0xCC))
  $tb.BorderThickness = "1"
  $tb.Padding = "8,0"
  return $tb
}

# Section label (muted uppercase)
function New-DlgSectionLabel {
  param([string]$Text)
  $tb = New-Object System.Windows.Controls.TextBlock
  $tb.Text       = $Text.ToUpper()
  $tb.FontSize   = 10
  $tb.FontWeight = "SemiBold"
  $tb.Foreground = [System.Windows.Media.SolidColorBrush]::new(
    [System.Windows.Media.Color]::FromRgb(0xAA,0xAA,0xAA))
  $tb.Margin     = "0,0,0,4"
  return $tb
}

# Dialog footer bar (border + button row)
function New-DlgFooter {
  param([System.Windows.Controls.Button[]]$Buttons)
  $border = New-Object System.Windows.Controls.Border
  $border.BorderBrush = [System.Windows.Media.SolidColorBrush]::new(
    [System.Windows.Media.Color]::FromRgb(0xE8,0xE8,0xE8))
  $border.BorderThickness = "0,1,0,0"
  $border.Padding = "14,10,14,10"
  $border.Background = [System.Windows.Media.SolidColorBrush]::new(
    [System.Windows.Media.Color]::FromRgb(0xFA,0xFA,0xFA))

  $panel = New-Object System.Windows.Controls.StackPanel
  $panel.Orientation = "Horizontal"
  $panel.HorizontalAlignment = "Right"
  foreach ($b in $Buttons) {
    $b.Margin = "8,0,0,0"
    $panel.Children.Add($b) | Out-Null
  }
  $border.Child = $panel
  return $border
}
# ───────────────────────────────────────────────────────────────────────────────

function Show-SettingsDialog {
  $dlg = New-StyledDialog -Title "Settings" -Width 800 -Height 330 -NoResize

  # Outer DockPanel: footer pinned to bottom
  $outer = New-Object System.Windows.Controls.DockPanel
  $outer.LastChildFill = $true

  # ── Footer ──────────────────────────────────────────────────────────────────
  $btnDefaults = New-DlgButton "Defaults"     -Width 90
  $btnSave     = New-DlgButton "Save"         -Width 90 -Primary
  $btnCancel   = New-DlgButton "Cancel"       -Width 90
  $footer = New-DlgFooter -Buttons @($btnDefaults, $btnSave, $btnCancel)
  [System.Windows.Controls.DockPanel]::SetDock($footer, "Bottom")
  $outer.Children.Add($footer) | Out-Null

  # ── Content area ─────────────────────────────────────────────────────────────
  $content = New-Object System.Windows.Controls.StackPanel
  $content.Margin = "16,14,16,10"

  # Hint text
  $hint = New-Object System.Windows.Controls.TextBlock
  $hint.Text = "If you don't set custom paths, everything stays under %ProgramData%\karanik_WinMaintenance."
  $hint.TextWrapping = "Wrap"
  $hint.FontSize   = 11
  $hint.Foreground = [System.Windows.Media.SolidColorBrush]::new(
    [System.Windows.Media.Color]::FromRgb(0x88,0x88,0x88))
  $hint.Margin = "0,0,0,12"
  $content.Children.Add($hint) | Out-Null

  # Row builder: label | textbox | Browse | Go
  $rows = @{}
  $rowDefs = @(
    @{ Key="Base";  Label="Base folder";  Value=$State.Paths.Base  },
    @{ Key="Logs";  Label="Logs folder";  Value=$State.Paths.Logs  },
    @{ Key="Cache"; Label="Cache folder"; Value=$State.Paths.Cache },
    @{ Key="Temp";  Label="Temp folder";  Value=$State.Paths.Temp  }
  )

  foreach ($rd in $rowDefs) {
    $row = New-Object System.Windows.Controls.Grid
    $row.Margin = "0,0,0,8"

    $c0 = New-Object System.Windows.Controls.ColumnDefinition; $c0.Width = "110"; $row.ColumnDefinitions.Add($c0) | Out-Null
    $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = "*";   $row.ColumnDefinitions.Add($c1) | Out-Null
    $c2 = New-Object System.Windows.Controls.ColumnDefinition; $c2.Width = "80";  $row.ColumnDefinitions.Add($c2) | Out-Null
    $c3 = New-Object System.Windows.Controls.ColumnDefinition; $c3.Width = "46";  $row.ColumnDefinitions.Add($c3) | Out-Null
    $c4 = New-Object System.Windows.Controls.ColumnDefinition; $c4.Width = "60";  $row.ColumnDefinitions.Add($c4) | Out-Null

    $lbl = New-Object System.Windows.Controls.TextBlock
    $lbl.Text = $rd.Label
    $lbl.VerticalAlignment = "Center"
    $lbl.FontSize = 12
    $lbl.Foreground = [System.Windows.Media.SolidColorBrush]::new(
      [System.Windows.Media.Color]::FromRgb(0x33,0x33,0x33))
    [System.Windows.Controls.Grid]::SetColumn($lbl, 0)

    $tb = New-DlgTextBox -Value $rd.Value -Height 26
    $tb.Margin = "0,0,6,0"
    [System.Windows.Controls.Grid]::SetColumn($tb, 1)

    $btnB = New-DlgButton "Browse..." -Width 74 -Height 26
    $btnB.Margin = "0,0,4,0"
    [System.Windows.Controls.Grid]::SetColumn($btnB, 2)

    $btnG = New-DlgButton "Go" -Width 40 -Height 26
    $btnG.ToolTip = "Open in Explorer"
    [System.Windows.Controls.Grid]::SetColumn($btnG, 3)

    # Clean button only for Logs, Cache, Temp (not Base)
    $btnC = $null
    if ($rd.Key -ne "Base") {
        $btnC = New-DlgButton "Clean" -Width 54 -Height 26
        $btnC.Margin = "4,0,0,0"
        $btnC.ToolTip = ("Delete all contents of {0} folder" -f $rd.Label)
        $btnC.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xC0,0x20,0x20))
        [System.Windows.Controls.Grid]::SetColumn($btnC, 4)
        $row.Children.Add($btnC) | Out-Null
    }

    $row.Children.Add($lbl)  | Out-Null
    $row.Children.Add($tb)   | Out-Null
    $row.Children.Add($btnB) | Out-Null
    $row.Children.Add($btnG) | Out-Null

    $content.Children.Add($row) | Out-Null
    $rows[$rd.Key] = @{ TextBox=$tb; BrowseBtn=$btnB; GoBtn=$btnG; CleanBtn=$btnC }
  }

  $outer.Children.Add($content) | Out-Null
  $dlg.Content = $outer

  function Browse-Folder([System.Windows.Controls.TextBox]$tb){
    Add-Type -AssemblyName System.Windows.Forms
    $f = New-Object System.Windows.Forms.FolderBrowserDialog
    $f.Description = "Select folder"
    $f.SelectedPath = if ($tb.Text.Trim()) { $tb.Text.Trim() } else { $env:ProgramData }
    if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $tb.Text = $f.SelectedPath }
  }

  $rows["Base"].BrowseBtn.Add_Click({  Browse-Folder $rows["Base"].TextBox  })
  $rows["Logs"].BrowseBtn.Add_Click({  Browse-Folder $rows["Logs"].TextBox  })
  $rows["Cache"].BrowseBtn.Add_Click({ Browse-Folder $rows["Cache"].TextBox })
  $rows["Temp"].BrowseBtn.Add_Click({  Browse-Folder $rows["Temp"].TextBox  })

  $rows["Base"].GoBtn.Add_Click({  Open-FolderInExplorer $rows["Base"].TextBox.Text  })
  $rows["Logs"].GoBtn.Add_Click({  Open-FolderInExplorer $rows["Logs"].TextBox.Text  })
  $rows["Cache"].GoBtn.Add_Click({ Open-FolderInExplorer $rows["Cache"].TextBox.Text })
  $rows["Temp"].GoBtn.Add_Click({  Open-FolderInExplorer $rows["Temp"].TextBox.Text  })

  function Clean-Folder {
    param([string]$Path, [string]$Label)
    if (-not $Path -or -not (Test-Path $Path)) { [System.Windows.MessageBox]::Show("Folder not found: $Path","Clean") | Out-Null; return }
    $r = [System.Windows.MessageBox]::Show(
      ("Delete ALL contents of:`n{0}`n`nThis cannot be undone." -f $Path),
      ("Clean {0}" -f $Label),
      [System.Windows.MessageBoxButton]::YesNo,
      [System.Windows.MessageBoxImage]::Warning)
    if ($r -ne [System.Windows.MessageBoxResult]::Yes) { return }
    $items = @(Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue)
    $ok = 0; $fail = 0
    foreach ($item in $items) {
      try { Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop; $ok++ }
      catch { $fail++ }
    }
    $msg = "Deleted $ok item(s)." + $(if ($fail -gt 0) { " Failed: $fail." } else { "" })
    Ui-Log ("Clean {0}: {1}" -f $Label, $msg) "INFO"
    [System.Windows.MessageBox]::Show($msg, ("Clean {0} done" -f $Label)) | Out-Null
  }

  $rows["Logs"].CleanBtn.Add_Click({  Clean-Folder $rows["Logs"].TextBox.Text  "Logs"  })
  $rows["Cache"].CleanBtn.Add_Click({ Clean-Folder $rows["Cache"].TextBox.Text "Cache" })
  $rows["Temp"].CleanBtn.Add_Click({  Clean-Folder $rows["Temp"].TextBox.Text  "Temp"  })

  $btnDefaults.Add_Click({
    Reset-ToDefaults
    $rows["Base"].TextBox.Text  = $State.Paths.Base
    $rows["Logs"].TextBox.Text  = $State.Paths.Logs
    $rows["Cache"].TextBox.Text = $State.Paths.Cache
    $rows["Temp"].TextBox.Text  = $State.Paths.Temp
  })

  $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })

  $btnSave.Add_Click({
    $State.Paths.Base  = $rows["Base"].TextBox.Text.Trim()
    $State.Paths.Logs  = $rows["Logs"].TextBox.Text.Trim()
    $State.Paths.Cache = $rows["Cache"].TextBox.Text.Trim()
    $State.Paths.Temp  = $rows["Temp"].TextBox.Text.Trim()

    if (-not $State.Paths.Base)  { $State.Paths.Base  = $DefaultBaseDir }
    if (-not $State.Paths.Logs)  { $State.Paths.Logs  = (Join-Path $State.Paths.Base "Logs") }
    if (-not $State.Paths.Cache) { $State.Paths.Cache = (Join-Path $State.Paths.Base "Cache") }
    if (-not $State.Paths.Temp)  { $State.Paths.Temp  = (Join-Path $State.Paths.Base "Temp") }

    Ensure-Dirs -Paths $State.Paths
    Save-Config
    $dlg.DialogResult = $true
    $dlg.Close()
  })

  $dlg.ShowDialog() | Out-Null
}

function Show-FixToolsPicker {
  $dlg = New-StyledDialog -Title "Fix Tools" -Width 500 -Height 480

  $outer = New-Object System.Windows.Controls.DockPanel
  $outer.LastChildFill = $true

  # ── Footer ──
  $btnRunSel = New-DlgButton "Run Selected" -Width 110 -Primary
  $btnCancel = New-DlgButton "Cancel" -Width 90
  $footer = New-DlgFooter -Buttons @($btnRunSel, $btnCancel)
  [System.Windows.Controls.DockPanel]::SetDock($footer, "Bottom")
  $outer.Children.Add($footer) | Out-Null

  # ── Content ──
  $body = New-Object System.Windows.Controls.DockPanel
  $body.Margin = "16,12,16,12"
  $body.LastChildFill = $true

  $hdr = New-Object System.Windows.Controls.StackPanel
  $hdr.Margin = "0,0,0,8"
  [System.Windows.Controls.DockPanel]::SetDock($hdr, "Top")

  $lbl = New-Object System.Windows.Controls.TextBlock
  $lbl.Text = "Select items to run (multi-select with Ctrl/Shift)"
  $lbl.FontSize = 12
  $lbl.Foreground = [System.Windows.Media.SolidColorBrush]::new(
    [System.Windows.Media.Color]::FromRgb(0x44,0x44,0x44))
  $hdr.Children.Add($lbl) | Out-Null
  $body.Children.Add($hdr) | Out-Null

  $lst = New-DlgListBox -SelectionMode "Extended"

  foreach ($i in ($FixCore | Sort-Object { [double]($_.Id -replace '[^0-9.]','') })) {
    $li = New-Object System.Windows.Controls.ListBoxItem
    $li.Content = (Strip-MenuPrefix $i.Title)
    $li.Padding = "8,5"
    $li.Tag = $i
    $lst.Items.Add($li) | Out-Null
  }

  # Teams section divider
  $sepLbl = New-Object System.Windows.Controls.ListBoxItem
  $sepLbl.Content = "TEAMS MAINTENANCE"
  $sepLbl.IsEnabled = $false
  $sepLbl.FontSize = 10
  $sepLbl.FontWeight = "SemiBold"
  $sepLbl.Foreground = [System.Windows.Media.SolidColorBrush]::new(
    [System.Windows.Media.Color]::FromRgb(0xAA,0xAA,0xAA))
  $sepLbl.Padding = "8,8,8,4"
  $sepLbl.Margin = "0,4,0,0"
  $lst.Items.Add($sepLbl) | Out-Null

  foreach ($i in ($FixTeams | Sort-Object { [double]($_.Id -replace '[^0-9.]','') })) {
    $li = New-Object System.Windows.Controls.ListBoxItem
    $li.Content = (Strip-MenuPrefix $i.Title)
    $li.Padding = "20,5,8,5"
    $li.Tag = $i
    $lst.Items.Add($li) | Out-Null
  }

  $body.Children.Add($lst) | Out-Null
  $outer.Children.Add($body) | Out-Null
  $dlg.Content = $outer

  $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })

  $picked = $null
  $btnRunSel.Add_Click({
    $sel = @($lst.SelectedItems | Where-Object { $_.Tag } | ForEach-Object { $_.Tag })
    if ($sel.Count -lt 1) {
      [System.Windows.MessageBox]::Show("Select at least one item.", $AppName) | Out-Null
      return
    }
    $script:picked = $sel
    $dlg.DialogResult = $true
    $dlg.Close()
  })

  $res = $dlg.ShowDialog()
  if ($res -ne $true) { return }
  Start-RunBatch -Items $script:picked
}

function Show-SerialRunDialog {
  param(
    [pscustomobject[]]$PreSelected = @()   # items to pre-populate in the right list
  )
  $all = @(Get-AllRunnableCatalogItems | Sort-Object Title)

  $dlg = New-StyledDialog -Title "Run Serial - Build sequence" -Width 1000 -Height 540

  # Outer DockPanel: footer at bottom, content fills rest
  $outer = New-Object System.Windows.Controls.DockPanel
  $outer.LastChildFill = $true

  # ── Footer ──────────────────────────────────────────────────────────────────
  $hintTxt = New-Object System.Windows.Controls.TextBlock
  $hintTxt.Text = "Tip: Ctrl/Shift to multi-select on the left, then Add. Double-click to add/remove."
  $hintTxt.FontSize = 11
  $hintTxt.VerticalAlignment = "Center"
  $hintTxt.Foreground = [System.Windows.Media.SolidColorBrush]::new(
    [System.Windows.Media.Color]::FromRgb(0x88,0x88,0x88))

  $btnRunSeq = New-DlgButton "Run Serial" -Width 110 -Primary
  $btnCancel = New-DlgButton "Cancel" -Width 90

  $footerBorder = New-Object System.Windows.Controls.Border
  $footerBorder.BorderBrush = [System.Windows.Media.SolidColorBrush]::new(
    [System.Windows.Media.Color]::FromRgb(0xE8,0xE8,0xE8))
  $footerBorder.BorderThickness = "0,1,0,0"
  $footerBorder.Background = [System.Windows.Media.SolidColorBrush]::new(
    [System.Windows.Media.Color]::FromRgb(0xFA,0xFA,0xFA))
  $footerBorder.Padding = "14,10,14,10"

  $footerDock = New-Object System.Windows.Controls.DockPanel
  $footerDock.LastChildFill = $false
  [System.Windows.Controls.DockPanel]::SetDock($hintTxt, "Left")
  $footerDock.Children.Add($hintTxt) | Out-Null

  $btnRow = New-Object System.Windows.Controls.StackPanel
  $btnRow.Orientation = "Horizontal"
  [System.Windows.Controls.DockPanel]::SetDock($btnRow, "Right")
  $btnRunSeq.Margin = "0,0,0,0"
  $btnCancel.Margin = "8,0,0,0"
  $btnRow.Children.Add($btnRunSeq) | Out-Null
  $btnRow.Children.Add($btnCancel) | Out-Null
  $footerDock.Children.Add($btnRow) | Out-Null

  $footerBorder.Child = $footerDock
  [System.Windows.Controls.DockPanel]::SetDock($footerBorder, "Bottom")
  $outer.Children.Add($footerBorder) | Out-Null

  # ── Main content ─────────────────────────────────────────────────────────────
  $body = New-Object System.Windows.Controls.Grid
  $body.Margin = "14,12,14,12"

  $c0 = New-Object System.Windows.Controls.ColumnDefinition; $c0.Width = "*";    $body.ColumnDefinitions.Add($c0) | Out-Null
  $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = "Auto"; $body.ColumnDefinitions.Add($c1) | Out-Null
  $c2 = New-Object System.Windows.Controls.ColumnDefinition; $c2.Width = "*";    $body.ColumnDefinitions.Add($c2) | Out-Null

  $r0 = New-Object System.Windows.Controls.RowDefinition; $r0.Height = "Auto"; $body.RowDefinitions.Add($r0) | Out-Null
  $r1 = New-Object System.Windows.Controls.RowDefinition; $r1.Height = "*";    $body.RowDefinitions.Add($r1) | Out-Null

  # Column headers
  $lblL = New-Object System.Windows.Controls.TextBlock
  $lblL.Text = "AVAILABLE ACTIONS"
  $lblL.FontSize = 10; $lblL.FontWeight = "SemiBold"
  $lblL.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xAA,0xAA,0xAA))
  $lblL.Margin = "0,0,0,6"
  [System.Windows.Controls.Grid]::SetRow($lblL, 0); [System.Windows.Controls.Grid]::SetColumn($lblL, 0)
  $body.Children.Add($lblL) | Out-Null

  $lblR = New-Object System.Windows.Controls.TextBlock
  $lblR.Text = "SELECTED SEQUENCE (top to bottom)"
  $lblR.FontSize = 10; $lblR.FontWeight = "SemiBold"
  $lblR.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xAA,0xAA,0xAA))
  $lblR.Margin = "0,0,0,6"
  [System.Windows.Controls.Grid]::SetRow($lblR, 0); [System.Windows.Controls.Grid]::SetColumn($lblR, 2)
  $body.Children.Add($lblR) | Out-Null

  # Left listbox
  $lstAvailable = New-DlgListBox -SelectionMode "Extended"
  foreach ($i in $all) {
    $li = New-Object System.Windows.Controls.ListBoxItem
    $li.Content = (Strip-MenuPrefix $i.Title)
    $li.Padding = "8,5"
    $li.Tag = $i
    $lstAvailable.Items.Add($li) | Out-Null
  }
  [System.Windows.Controls.Grid]::SetRow($lstAvailable, 1); [System.Windows.Controls.Grid]::SetColumn($lstAvailable, 0)
  $body.Children.Add($lstAvailable) | Out-Null

  # Right listbox
  $lstSelected = New-DlgListBox -SelectionMode "Single"
  [System.Windows.Controls.Grid]::SetRow($lstSelected, 1); [System.Windows.Controls.Grid]::SetColumn($lstSelected, 2)
  $body.Children.Add($lstSelected) | Out-Null

  # Pre-populate from PreSelected parameter (e.g. from "Add to Run Serial" context menu)
  foreach ($pre in $PreSelected) {
    $pli = New-Object System.Windows.Controls.ListBoxItem
    $pli.Content = (Strip-MenuPrefix $pre.Title)
    $pli.Padding = "8,5"
    $pli.Tag = $pre
    $lstSelected.Items.Add($pli) | Out-Null
  }
  if ($lstSelected.Items.Count -gt 0) { $lstSelected.SelectedIndex = $lstSelected.Items.Count - 1 }

  # Middle buttons
  $mid = New-Object System.Windows.Controls.StackPanel
  $mid.Orientation = "Vertical"
  $mid.VerticalAlignment = "Center"
  $mid.Margin = "10,0,10,0"
  [System.Windows.Controls.Grid]::SetRow($mid, 1); [System.Windows.Controls.Grid]::SetColumn($mid, 1)

  $btnAdd    = New-DlgButton "Add -->"    -Width 94 -Height 28 -Primary; $btnAdd.Margin    = "0,0,0,6"
  $btnRemove = New-DlgButton "<-- Remove" -Width 94 -Height 28; $btnRemove.Margin = "0,0,0,18"
  $btnUp     = New-DlgButton "Up"         -Width 94 -Height 28 -Primary; $btnUp.Margin     = "0,0,0,6"
  $btnDown   = New-DlgButton "Down"       -Width 94 -Height 28 -Primary

  $mid.Children.Add($btnAdd)    | Out-Null
  $mid.Children.Add($btnRemove) | Out-Null
  $mid.Children.Add($btnUp)     | Out-Null
  $mid.Children.Add($btnDown)   | Out-Null
  $body.Children.Add($mid) | Out-Null

  $outer.Children.Add($body) | Out-Null
  $dlg.Content = $outer

  function Add-ItemsToSelected {
    $items = @($lstAvailable.SelectedItems)
    if ($items.Count -lt 1) { return }
    foreach ($li in $items) {
      $new = New-Object System.Windows.Controls.ListBoxItem
      $new.Content = $li.Content
      $new.Padding = "8,5"
      $new.Tag = $li.Tag
      $lstSelected.Items.Add($new) | Out-Null
    }
    if ($lstSelected.Items.Count -gt 0 -and -not $lstSelected.SelectedItem) {
      $lstSelected.SelectedIndex = $lstSelected.Items.Count - 1
    }
  }

  function Remove-SelectedItem {
    $sel = $lstSelected.SelectedItem
    if (-not $sel) { return }
    $idx = $lstSelected.SelectedIndex
    $lstSelected.Items.Remove($sel)
    if ($lstSelected.Items.Count -gt 0) {
      $lstSelected.SelectedIndex = [Math]::Min($idx, $lstSelected.Items.Count - 1)
    }
  }

  function Move-Selected([int]$delta) {
    $idx = $lstSelected.SelectedIndex
    if ($idx -lt 0) { return }
    $newIdx = $idx + $delta
    if ($newIdx -lt 0 -or $newIdx -ge $lstSelected.Items.Count) { return }
    $item = $lstSelected.Items[$idx]
    $lstSelected.Items.RemoveAt($idx)
    $lstSelected.Items.Insert($newIdx, $item)
    $lstSelected.SelectedIndex = $newIdx
  }

  $btnAdd.Add_Click({ Add-ItemsToSelected })
  $btnRemove.Add_Click({ Remove-SelectedItem })
  $btnUp.Add_Click({ Move-Selected -1 })
  $btnDown.Add_Click({ Move-Selected 1 })

  $lstAvailable.Add_MouseDoubleClick({ Add-ItemsToSelected })
  $lstSelected.Add_MouseDoubleClick({ Remove-SelectedItem })

  $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })

  $selectedCatalogItems = $null
  $btnRunSeq.Add_Click({
    if ($lstSelected.Items.Count -lt 1) {
      [System.Windows.MessageBox]::Show("Add at least one action to the sequence.", $AppName) | Out-Null
      return
    }
    $script:selectedCatalogItems = @($lstSelected.Items | ForEach-Object { $_.Tag })
    $dlg.DialogResult = $true
    $dlg.Close()
  })

  $res = $dlg.ShowDialog()
  if ($res -ne $true) { return $null }
  return $script:selectedCatalogItems
}
#endregion Dialogs

#region Menu events

# ── File menu ────────────────────────────────────────────────────────────────
$mnuExit.Add_Click({
    # Close() triggers the Window.Closing event, so Stealth Mode cleanup runs here too
    $Window.Close()
})
$mnuReload.Add_Click({ Build-Tree; Ui-Log "Menu reloaded." "INFO" })
$mnuOpenLog.Add_Click({
  try { Open-LatestLog }
  catch { Ui-Append (Write-LogLine -Message (Format-ErrorFull $_) -Level "ERROR") }
})
$mnuSettings.Add_Click({
  try { Show-SettingsDialog }
  catch { Ui-Append (Write-LogLine -Message (Format-ErrorFull $_) -Level "ERROR") }
})

# ── Tools menu checkboxes kept in sync with toolbar toggles ─────────────────
$mnuVerbose.Add_Checked({
  $State.VerboseOutput = $true
  $tglVerbose.IsChecked = $true
  Save-Config
  Ui-Log "Verbose Output enabled." "INFO"
})
$mnuVerbose.Add_Unchecked({
  $State.VerboseOutput = $false
  $tglVerbose.IsChecked = $false
  Save-Config
  Ui-Log "Verbose Output disabled." "INFO"
})

# ── Run action helper (shared by menu + toolbar button) ──────────────────────
function Invoke-RunSelected {
  try {
    $item = Get-SelectedLeafItem
    if (-not $item) { Ui-Log "Select an action from the tree first." "WARN"; return }
    if ($item.Type -eq "Picker") { Show-FixToolsPicker; return }
    Start-RunRemoteItem -Item $item
  } catch {
    Ui-Append (Write-LogLine -Message (Format-ErrorFull $_) -Level "ERROR")
    Ui-SetStatus "Error."
    Set-UiBusy $false
  }
}

$mnuRun.Add_Click({ Invoke-RunSelected })
$btnRun.Add_Click({ Invoke-RunSelected })

# ── Run Serial ───────────────────────────────────────────────────────────────
function Invoke-RunSerial {
  try {
    # Pre-populate with any items queued via right-click "Add to Run Serial"
    $preItems = if ($Script:SerialQueue -and $Script:SerialQueue.Count -gt 0) {
      $Script:SerialQueue.ToArray()
    } else { @() }
    $Script:SerialQueue = $null   # clear queue once dialog opens

    $items = Show-SerialRunDialog -PreSelected $preItems
    if (-not $items) { return }
    Start-RunBatch -Items $items
  } catch {
    Ui-Append (Write-LogLine -Message (Format-ErrorFull $_) -Level "ERROR")
    Ui-SetStatus "Error."
    Set-UiBusy $false
  }
}

$mnuRunSerial.Add_Click({ Invoke-RunSerial })
$btnRunSerial.Add_Click({ Invoke-RunSerial })

# ── Toolbar ToggleButtons sync with menu checkboxes ──────────────────────────
$tglForce.Add_Checked({
  $mnuForce.IsChecked = $true
  Ui-Log "Force re-download enabled." "INFO"
})
$tglForce.Add_Unchecked({
  $mnuForce.IsChecked = $false
  Ui-Log "Force re-download disabled." "INFO"
})

$tglVerbose.Add_Checked({
  $State.VerboseOutput = $true
  $mnuVerbose.IsChecked = $true
  Save-Config
  Ui-Log "Verbose Output enabled." "INFO"
})
$tglVerbose.Add_Unchecked({
  $State.VerboseOutput = $false
  $mnuVerbose.IsChecked = $false
  Save-Config
  Ui-Log "Verbose Output disabled." "INFO"
})

# ── Clear log button ──────────────────────────────────────────────────────────
$btnClearLog.Add_Click({ Ui-ClearLog; Ui-Log "Log cleared." "INFO" })
$btnClearTerminal.Add_Click({ Ui-ClearTerminal })

# ── Filter textbox  -  rebuild tree on each keystroke ──────────────────────────
$txtFilter.Add_TextChanged({
  Build-Tree -Filter $txtFilter.Text
})

# ── Theme dropdown ────────────────────────────────────────────────────────────
$btnTheme.Add_Click({ $popTheme.IsOpen = -not $popTheme.IsOpen })
$btnClean.Add_Click({ $popClean.IsOpen = -not $popClean.IsOpen })

function Clean-FolderQuick {
  param([string]$Path, [string]$Label)
  if (-not $Path -or -not (Test-Path $Path)) {
    Ui-Log ("Clean {0}: folder not found ({1})" -f $Label, $Path) "WARN"; return
  }
  $r = [System.Windows.MessageBox]::Show(
    ("Delete ALL contents of:`n{0}`n`nThis cannot be undone." -f $Path),
    ("Clean {0}" -f $Label),
    [System.Windows.MessageBoxButton]::YesNo,
    [System.Windows.MessageBoxImage]::Warning)
  if ($r -ne [System.Windows.MessageBoxResult]::Yes) { return }
  $items = @(Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue)
  $ok = 0; $fail = 0
  foreach ($item in $items) {
    try { Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop; $ok++ }
    catch { $fail++ }
  }
  $msg = "Cleaned {0}: {1} item(s) deleted." -f $Label, $ok
  if ($fail -gt 0) { $msg += " ($fail failed)" }
  Ui-Log $msg "SUCCESS"
}

$btnCleanLogs.Add_Click({
  $popClean.IsOpen = $false
  Clean-FolderQuick $State.Paths.Logs "Logs"
})
$btnCleanCache.Add_Click({
  $popClean.IsOpen = $false
  Clean-FolderQuick $State.Paths.Cache "Cache"
})
$btnCleanTemp.Add_Click({
  $popClean.IsOpen = $false
  Clean-FolderQuick $State.Paths.Temp "Temp"
})

function Apply-Theme {
  param([string]$Mode)   # "Auto", "Light", "Dark"
  try {
    switch ($Mode) {
      "Light" {
        $Window.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xFF,0xFF,0xFF))
        $Window.Foreground = [System.Windows.Media.Brushes]::Black
        $txtOutput.Foreground   = [System.Windows.Media.Brushes]::Black
        $txtOutput.Background   = [System.Windows.Media.Brushes]::White
        $txtTerminal.Foreground = [System.Windows.Media.Brushes]::Black
        $txtTerminal.Background = [System.Windows.Media.Brushes]::White
        $tvMenu.Background      = [System.Windows.Media.Brushes]::Transparent
        $tvMenu.Foreground      = [System.Windows.Media.Brushes]::Black
        # Reset all TreeViewItems to light colors
        foreach ($item in $tvMenu.Items) {
            $item.Foreground = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x55,0x55,0x55))
            $item.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xF2,0xF4,0xF7))
            foreach ($leaf in $item.Items) { $leaf.Foreground = [System.Windows.Media.Brushes]::Black }
        }
      }
      "Dark" {
        $darkBg   = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x1E,0x1E,0x1E))
        $darkFg   = [System.Windows.Media.Brushes]::White
        $darkGrpBg= [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0x2D,0x2D,0x2D))
        $darkGrpFg= [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xBB,0xBB,0xBB))
        $darkLeafFg = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0xE8,0xE8,0xE8))

        $Window.Background      = $darkBg
        $Window.Foreground      = $darkFg
        $txtOutput.Foreground   = [System.Windows.Media.Brushes]::LightGray
        $txtOutput.Background   = $darkBg
        $txtTerminal.Foreground = [System.Windows.Media.Brushes]::LightGray
        $txtTerminal.Background = $darkBg
        $tvMenu.Background      = $darkBg
        $tvMenu.Foreground      = $darkFg
        # Apply dark colors to all TreeViewItems (groups + leaves)
        foreach ($item in $tvMenu.Items) {
            $item.Foreground = $darkGrpFg
            $item.Background = $darkGrpBg
            foreach ($leaf in $item.Items) { $leaf.Foreground = $darkLeafFg }
        }
      }
      default {
        $Window.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
        $Window.ClearValue([System.Windows.Controls.Control]::ForegroundProperty)
        $txtOutput.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
        $txtOutput.ClearValue([System.Windows.Controls.Control]::ForegroundProperty)
        $txtTerminal.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
        $txtTerminal.ClearValue([System.Windows.Controls.Control]::ForegroundProperty)
        $tvMenu.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
        $tvMenu.ClearValue([System.Windows.Controls.Control]::ForegroundProperty)
        foreach ($item in $tvMenu.Items) {
            $item.ClearValue([System.Windows.Controls.Control]::ForegroundProperty)
            $item.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
            foreach ($leaf in $item.Items) { $leaf.ClearValue([System.Windows.Controls.Control]::ForegroundProperty) }
        }
      }
    }
    Ui-Log ("Theme set to: {0}" -f $Mode) "INFO"
    $State.Theme = $Mode
    Save-Config
  } catch { }
}

$btnThemeAuto.Add_Click({
  $popTheme.IsOpen = $false
  Apply-Theme "Auto"
})
$btnThemeLight.Add_Click({
  $popTheme.IsOpen = $false
  Apply-Theme "Light"
})
$btnThemeDark.Add_Click({
  $popTheme.IsOpen = $false
  Apply-Theme "Dark"
})

# ── About ─────────────────────────────────────────────────────────────────────
$mnuAbout.Add_Click({ try { Start-Process $AboutUrl | Out-Null } catch { } })

#endregion Menu events

# ── Stealth Mode: wipe everything on close ────────────────────────────────────
# Store reference at script scope so the Closing event can access them
$Script:_stealthToggle = $tglStealth
$Script:_stealthPath   = $State.Paths.Base

$Window.Add_Closing({
    if ($Script:_stealthToggle.IsChecked -eq $true) {
        $basePath = $Script:_stealthPath
        if (-not $basePath) { $basePath = "C:\ProgramData\karanik_WinMaintenance" }
        if (Test-Path $basePath) {
            try {
                [System.IO.Directory]::Delete($basePath, $true)
            } catch {
                # Fallback: use cmd.exe rd which handles locked files better
                & cmd.exe /c "rd /s /q `"$basePath`"" 2>$null
            }
        }
    }
})

$null = $Window.ShowDialog()
