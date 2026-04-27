# karanik_WinMaintenance

[![GitHub release](https://img.shields.io/badge/version-1.7-blue?style=flat-square)](https://github.com/karanikn/karanik_WinMaintenance)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207.x-blue?style=flat-square&logo=powershell)](https://github.com/PowerShell/PowerShell)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11%20%7C%20Server-lightgrey?style=flat-square&logo=windows)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/license-GPL--3.0-blue?style=flat-square)](LICENSE)
[![AI Assisted](https://img.shields.io/badge/built%20with-Claude%20AI-orange?style=flat-square&logo=anthropic)](https://claude.ai)

> **A powerful WPF-based PowerShell launcher for Windows system administration and maintenance.**  
> Download → Cache → Run. Clean, fast, and always up to date.

---

## 📸 Screenshots

| Main Window | Run Serial |
|---|---|
| ![Main Window](https://raw.githubusercontent.com/karanikn/karanik_WinMaintenance_Online/main/Screenshots/karanik_WinMaintenance.png) | ![Run Serial](https://raw.githubusercontent.com/karanikn/karanik_WinMaintenance_Online/main/Screenshots/karanik_WinMaintenanceRunSerial.png) |

| PowerShell Output | Settings |
|---|---|
| ![PowerShell Output](https://raw.githubusercontent.com/karanikn/karanik_WinMaintenance_Online/main/Screenshots/karanik_WinMaintenance_Powershell.png) | ![Settings](https://raw.githubusercontent.com/karanikn/karanik_WinMaintenance_Online/main/Screenshots/karanik_WinMaintenance_settings.png) |

| Office Manager | PS Module Manager |
|---|---|
| ![Office Manager](https://raw.githubusercontent.com/karanikn/karanik_WinMaintenance_Online/main/Screenshots/OfficeManager.png) | ![PS Module Manager](https://raw.githubusercontent.com/karanikn/karanik_WinMaintenance_Online/main/Screenshots/ModuleManager.png) |

| WinDiag-AI |
|---|
| ![WinDiag-AI](https://raw.githubusercontent.com/karanikn/WinDiag-AI/main/Screenshots/WinDiag-AI-Main.png) |

---

## ✨ Overview

**karanik_WinMaintenance** is a professional Windows maintenance toolkit built as a WPF GUI launcher written entirely in PowerShell. It provides a clean, organized interface for running a curated collection of system administration scripts — without needing to open a terminal or remember command syntax.

Scripts are **downloaded on demand** from a remote server, cached locally, and executed in an elevated PowerShell session. The result is a tool that is always up to date: updating a script on the server instantly propagates to all users on the next run.

Designed for **IT administrators, help desk engineers, and power users** who maintain Windows workstations and servers.

---

## 🚀 Quick Launch

### Option 1 — One-liner (no download needed)

Open **PowerShell as Administrator** and run:

```powershell
irm "https://karanik.gr/win" | iex
```

This downloads and launches the application in a single command — the same way tools like Chris Titus Tech's WinUtil work.

### Option 2 — Download and run the PS1

```powershell
# Download
Invoke-WebRequest -Uri "https://karanik.gr/scripts/powershell/karanik_WinMaintenance/karanik_WinMaintenance.ps1" -OutFile "$env:TEMP\karanik_WinMaintenance.ps1"

# Run as Administrator
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "$env:TEMP\karanik_WinMaintenance.ps1"
```

### Requirements

| Requirement | Details |
|---|---|
| OS | Windows 10 / Windows 11 / Windows Server 2016+ |
| PowerShell | 5.1 or later (PS7 also supported) |
| Privileges | Must run as **Administrator** |
| Internet | Required to download scripts on first run (cached after) |

---

## 🖥️ Interface

The application features a clean two-panel layout:

- **Left panel** — Collapsible script catalog organized by category
- **Right panel (top)** — Live log output with color-coded `[INFO]` / `[SUCCESS]` / `[WARN]` / `[ERROR]` entries
- **Right panel (bottom)** — Terminal output (raw stdout/stderr from child scripts)
- **Toolbar** — Run, Run Serial, Force re-download, Verbose, Stealth Mode
- **Search bar** — Real-time filtering of the script catalog
- **Theme selector** — Auto (follows system), Light, Dark
- **Settings** — Customizable paths for Base, Logs, Cache, and Temp folders

### Toolbar Buttons

| Button | Description |
|---|---|
| **Run ▶** | Runs the selected script(s) |
| **Run Serial...** | Runs all selected scripts sequentially, one after another |
| **Force re-download** | Bypasses cache and always fetches the latest version from the server |
| **Verbose** | Enables detailed debug output from scripts |
| **🕵 Stealth Mode** | When enabled, permanently deletes all application data on exit (`C:\ProgramData\karanik_WinMaintenance`) |

---

## 📋 Script Catalog

### 🔧 Main

| ID | Script | Description |
|---|---|---|
| 1.1 | Update Windows | Installs pending Windows Updates via PSWindowsUpdate + WUA COM fallback |
| 1.2 | Reset Windows Update | 5-step WU reset: stops services, clears SoftwareDistribution, restarts |
| 1.3 | Defender: Update signatures & Quick Scan | Updates Windows Defender definitions and runs a Quick Scan |
| 1.4 | GPUpdate /force | Forces Group Policy refresh for Computer and User |
| 1.5 | Create System Restore Point | Creates a named system restore point |
| 1.51 | Clean All Restore Points | Removes all existing restore points (with confirmation countdown) |
| 1.6 | Clear Event Logs | Clears all Windows Event Logs via wevtutil.exe |
| 1.7 | Install Winget | Installs or verifies the Windows Package Manager (App Installer) |
| 1.8 | Install latest PowerShell | WPF GUI showing PS5/PS7 versions vs latest; installs via winget or MSI |
| 1.9 | Upgrade all apps via Winget | WPF picker to selectively upgrade installed apps (excludes pinned) |

### ⚙️ PowerShell Tools

| ID | Script | Description |
|---|---|---|
| 2.1 | PS Module Manager | Full-featured PowerShell module manager (update, install, backup, restore) |

### 🧹 Cleaning Tools

| ID | Script | Description |
|---|---|---|
| 3.1 | Clean all Temporary files | Cleans system temp, all user profiles, WU download cache, font cache, browser caches, Recycle Bin, and runs cleanmgr |
| 3.2 | Clear Teams cache only | Stops Teams, clears all cache folders for all user profiles (Classic + MSIX) |
| 3.3 | Teams Full clean + reinstall | WPF GUI: selective uninstall steps + choice of Teams version to reinstall |
| 3.4 | Cleanup SoftwareDistribution | Clears SoftwareDistribution folder (preserves DataStore/history) |

### 📤 Export Tools

| ID | Script | Description |
|---|---|---|
| 4.1 | System Information | Full system report: OS, hardware, network, services, processes, event errors, remote tools. Saved as .txt and opened in Notepad |
| 4.2 | Export Installed Software | Registry-based software list exported to CSV on Desktop |
| 4.3 | Export Drivers | All installed drivers exported to CSV (Get-WindowsDriver + driverquery fallback) |
| 4.4 | Wi-Fi Profiles: Export | Exports all Wi-Fi profiles with clear-text keys to `Desktop\WiFi_Profiles\` |
| 4.5 | Wi-Fi Profiles: Import | Imports Wi-Fi profiles from `Desktop\WiFi_Profiles\` XML files |
| 4.6 | Scheduled Tasks Audit | All scheduled tasks exported to CSV; flags non-Ready tasks |
| 4.7 | Windows Update History | Collects update history from multiple sources (WUA, DISM, EventLog, HotFix) and exports to CSV; opens Settings → Update History |
| 4.8 | Battery Report | Generates powercfg battery report and opens it in the browser |

### 🔨 Fix Tools

| ID | Script | Description |
|---|---|---|
| 5.1 | Reset Print Spooler | Stops Spooler, clears spool files, restarts service |
| 5.2 | Network Reset | Resets Winsock, TCP/IP stack, releases/renews IP, flushes DNS |
| 5.3 | Wi-Fi Reset | Full network reset + removes stale dni_dne registry key |
| 5.4 | Set Classic Right-Click Menu | Enables the classic Windows 10-style context menu in Windows 11 |
| 5.5 | Restore Default Right-Click Menu | Restores the default Windows 11 context menu |
| 5.6 | Disable Hibernation | Runs `powercfg /h off` and verifies hiberfil.sys removal |
| 5.7 | Enable Hibernation | Runs `powercfg /h on` and verifies hiberfil.sys creation |
| 5.8 | WebView2 Repair / Install | Checks, downloads and installs Microsoft Edge WebView2 Runtime |
| 5.9 | WebView2 / Outlook Auth Repair (Advanced) | Full Outlook Modern Auth / WAM / AAD BrokerPlugin repair + diagnostics export |
| 5.10 | Remove iCloud (All Users) | Removes iCloud/Apple components from all user profiles; kills related processes |

### ⭐ Extra Tools

| ID | Tool | Description |
|---|---|---|
| 6.1 | RDS Licensing Grace Period Reset | Resets the RDS Licensing Grace Period on Windows Server — can be run as many times as needed |
| 6.2 | Microsoft Activation Scripts (MAS) | Launches `irm https://get.activated.win \| iex` in a new elevated window |
| 6.3 | WinScript | Launches `irm https://winscript.cc/irm \| iex` |
| 6.4 | Chris Titus Tech's Windows Utility | Launches `irm https://christitus.com/win \| iex` |
| 6.5 | WinDiag-AI | WPF GUI for Windows diagnostics with local AI analysis (Ollama); SMART, PDF export, driver check |

### 🏢 Office Tools

| ID | Script | Description |
|---|---|---|
| 7.1 | Office Tools | Click-to-Run update control, Quick/Online repair, SaRA Enterprise (13 scenarios) |

---

## 🗂️ File Structure

```
C:\ProgramData\karanik_WinMaintenance\
├── Logs\          ← Per-run log files (timestamped)
├── Cache\         ← Downloaded PS1 scripts
├── Temp\          ← Temporary working files
└── config.json    ← User settings (custom paths, theme)
```

All paths are fully configurable via **Settings**.

---

## 🔒 Stealth Mode

When **Stealth Mode** is enabled (🕵 toolbar button), closing the application **permanently deletes** the entire `C:\ProgramData\karanik_WinMaintenance` folder including all logs, cached scripts, and configuration. Files are not sent to the Recycle Bin.

Useful for deployments where no trace of the tool should remain after use.

---

## 🏗️ Architecture

| Component | Details |
|---|---|
| **Language** | PowerShell 5.1+ (compatible with both Windows PowerShell and PowerShell 7) |
| **UI** | WPF (Windows Presentation Foundation) via `[System.Windows.Markup.XamlReader]` |
| **Threading** | `[PowerShell]::Create()` + dedicated Runspace for background execution; `ConcurrentQueue` for thread-safe UI updates |
| **Script delivery** | `System.Net.WebClient.DownloadFile()` with validation (checks first 512 bytes for HTML/error responses) |
| **Logging** | Dual-stream: structured `[INFO]/[SUCCESS]/[WARN]/[ERROR]` to log file + raw stdout/stderr to terminal panel |
| **Execution types** | `Remote` (download+cache+run), `Standalone` (separate STA window for WPF scripts), `Inline` (irm\|iex in new elevated window) |

---

## 📝 Changelog

### v1.7 — April 2026

- **4 new scripts added to catalog:**
  - `5.8` **WebView2 Repair / Install** — checks, downloads and installs Microsoft Edge WebView2 Runtime; kills stale msedgewebview2.exe processes before install
  - `5.9` **WebView2 / Outlook Auth Repair (Advanced)** — comprehensive repair tool for Outlook Modern Auth / WAM / AAD BrokerPlugin issues; cleans caches, exports diagnostics, tests Microsoft 365 endpoints
  - `5.10` **Remove iCloud (All Users)** — removes iCloud/Apple components (processes, AppData folders, ProgramData) from all user profiles
  - `6.5` **WinDiag-AI** — WPF GUI for Windows diagnostics with local AI analysis via Ollama; includes SMART details, PDF export, driver check, dark/light theme (`Type=Standalone`)
- **Tooltips added to all 39 catalog entries** — each tooltip describes exact steps, affected locations, and use cases; written from script source analysis

### v1.6 — March 2026

- **Office Tools group added** (`7.1` Office Tools) — Click-to-Run update control (enable/disable/check/force), Quick and Online repair, Mail Control Panel, SaRA Enterprise with 13 scenarios (uninstall, Outlook scan, activation reset, Teams add-in fix, CalCheck)
- **PSModuleManager** — fixed 0-module scan issue inside launcher runspace; replaced inline `-Command` execution with temp `.ps1` + `.txt` output files via `System.Diagnostics.Process` to bypass stdout pipe truncation

### v1.5 — March 2026

- **Stealth Mode** — 🕵 toolbar toggle; on close permanently wipes `C:\ProgramData\karanik_WinMaintenance` using script-scope variable binding for reliable WPF Closing event access
- **Tooltips** — rich tooltips added to Run Serial, Force re-download, Verbose, and Stealth Mode buttons
- **URL fix** — corrected `karanik_WinMaintanance` → `karanik_WinMaintenance` typo in `$BaseUrl` and all absolute paths
- **One-liner launch** — `irm "https://karanik.gr/win" | iex` via Cloudflare route
- **SoftwareDistribution history protection** — ResetWindowsUpdate and Cleanup-SoftwareDistribution now exclude `DataStore` subfolder to preserve Windows Update history
- **PS5.1 compatibility** — fixed null-conditional `?.Source` operators (9 locations), 2-arg `Thickness::new()` constructors, `Clear-WinEvent` incompatibility throughout all scripts

### v1.4 — Fix Tools & Extra Tools

- **Fix Tools suite:** ResetPrintSpooler, NetworkReset, WiFiReset, SetClassicContextMenu, RestoreDefaultContextMenu, DisableHibernation, EnableHibernation
- **Extra Tools (Inline type):** Microsoft Activation Scripts, WinScript, Chris Titus Tech's Windows Utility — each launches in a new elevated PowerShell window
- **RDS Grace Period Manager** — full-screen Windows Server dashboard for RDS licensing management
- **Inline execution type** — new script type for `irm | iex` commands; bypasses download/cache pipeline entirely
- **Standalone execution type** — new script type for scripts with their own WPF GUI; runs in a separate STA window with environment variable passthrough via EncodedCommand wrapper

### v1.3 — Export Tools

- **SystemInfo** — comprehensive system report including OS, BIOS, CPU, RAM sticks, disk drives, network adapters, external IP lookup, netstat/ARP/routing, installed software, services, startup items, top processes by CPU, System/Application Event Log errors, hotfixes, TeamViewer/AnyDesk connection info, hosts file, chkdsk logs
- **ExportInstalledSoftware** — registry-based (HKLM + HKCU Uninstall keys), deduplicated, CSV to Desktop
- **ExportDrivers** — `Get-WindowsDriver` with `driverquery.exe` fallback, CSV to Desktop
- **ExportWiFiProfiles** — bulk `netsh wlan export` with `key=clear`; per-profile fallback if bulk returns 0 files
- **ImportWiFiProfiles** — import all XML profiles from `Desktop\WiFi_Profiles\`
- **ExportScheduledTasks** — full task audit with LastRunTime, NextRunTime, RunAs, Actions; flags non-Ready tasks
- **ExportWindowsUpdateHistory** — multi-source collection (WUA QueryHistory, Win32_ReliabilityRecords, DISM, System EventLog, Get-HotFix); exports CSV and opens Settings → Update History
- **BatteryReport** — `powercfg /batteryreport` to timestamped HTML on Desktop; skips gracefully on desktops with no battery

### v1.2 — Cleaning Tools

- **CleanTempFiles** — 7-step cleaner: Windows\Temp, Prefetch, all user profile temps, WU Download cache (stops/restarts WU services), font cache, browser caches (Chrome/Edge/Firefox/Teams), DNS flush, `cleanmgr /sagerun:100` with StateFlags0100, Windows.old removal with `takeown`/`icacls`, root drive junk files
- **ClearTeamsCache** — stops all Teams processes, clears per-user cache folders for both Classic Teams (`AppData\Roaming\Microsoft\Teams`) and New Teams MSIX (`Packages\MSTeams_*\LocalCache`)
- **TeamsCleanupAndInstall** — WPF GUI replacing all `Read-Host` prompts; checkboxes for 8 cleanup steps, radio buttons for Teams version to reinstall (New / Classic / MSI / cleanup only); runs steps immediately without countdown
- **Cleanup-SoftwareDistribution** — stops WU services, measures before/after size, clears SoftwareDistribution contents, removes `.old`/`.bak` backup folders, restarts services with dependency ordering

### v1.1 — Main Script Suite

- **UpdateWindows** — PSWindowsUpdate module with WUA COM fallback; transcript logging
- **ResetWindowsUpdate** — 5-step reset: stop services, clear SoftwareDistribution (excluding DataStore), re-register WU DLLs, flush DNS, restart services
- **DefenderUpdate** — `Update-MpSignature` + `Start-MpScan -ScanType QuickScan`
- **InstallPowerShell** — WPF GUI showing installed PS5/PS7 versions vs latest from GitHub API; installs via `winget` with MSI fallback; filters VT100 progress bar from output
- **UpgradeWingetApps** — WPF picker dialog; ANSI escape code stripping; ID-anchor regex parser for upgrade list; `--include-unknown` flag for packages with unknown current version
- **GPUpdate** — `gpupdate /force` for both Computer and User policy
- **CreateRestorePoint / CleanRestorePoints** — with countdown confirmation for destructive operation
- **ClearEventLogs** — `wevtutil.exe cl` for PS5.1 compatibility
- **InstallWinget** — checks existing installation, skips if already present
- **PSModuleManager** — full PowerShell module manager hosted separately

### v1.0 — Initial Release

- WPF GUI launcher with collapsible TreeView catalog, real-time search/filter
- Download → cache → run pipeline with response validation (HTML/error detection)
- Live log tailing from child script log files via `DispatcherTimer`
- Dual output panels: structured log (top 2/3) + terminal output (bottom 1/3) with draggable GridSplitter
- Background worker using `[PowerShell]::Create()` + `RunspaceFactory`; `ConcurrentQueue` for thread-safe UI updates
- Toolbar: Run ▶, Run Serial..., Force re-download, Verbose
- Theme selector: Auto (follows system) / Light / Dark
- Settings dialog: configurable Base, Logs, Cache, Temp paths with Browse/Go/Clean buttons
- Trash icon (🗑) with popup menu: Clean Logs / Clean Cache / Clean Temp
- Ctrl+click multi-select, double-click to run, right-click context menu
- Config persistence via `config.json`
- Status badge in toolbar showing running/completed/error state

---

## 👤 Author

**Nikolaos Karanikolas**  
🌐 [karanik.gr](https://karanik.gr)

---

## 🤖 AI Assistance

This project was developed with the assistance of **[Claude](https://claude.ai)** (Anthropic AI). The architecture, WPF GUI, threading model, async patterns, script execution pipeline, and all PowerShell code were designed and iterated collaboratively between the developer and Claude over an extended development session.

---

## ⚠️ Disclaimer

This tool executes PowerShell scripts with Administrator privileges. Always review scripts before running them in production environments. The author takes no responsibility for data loss or system damage resulting from use of this tool.
