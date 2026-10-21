# karanik_WinMaintenance

[![GitHub release](https://img.shields.io/badge/version-2.4-blue?style=flat-square)](https://github.com/karanikn/karanik_WinMaintenance)
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
| ![Office Manager](https://raw.githubusercontent.com/karanikn/karanik_WinMaintenance_Online/main/Screenshots/Office_Manager.png) | ![PS Module Manager](https://raw.githubusercontent.com/karanikn/karanik_WinMaintenance_Online/main/Screenshots/ModuleManager.png) |

| WinDiag-AI |
|---|
| ![WinDiag-AI](https://raw.githubusercontent.com/karanikn/WinDiag-AI/main/Screenshots/WinDiag-AI-Main.png) |

| Clean Temporary Files | DISM / SFC Repair |
|---|---|
| ![Clean Temporary Files](https://raw.githubusercontent.com/karanikn/karanik_WinMaintenance_Online/main/Screenshots/CleanTemp.png) | ![DISM / SFC Repair](https://raw.githubusercontent.com/karanikn/karanik_WinMaintenance_Online/main/Screenshots/DISM-SFC.png) |

| Repair-PrintSystem |
|---|
| ![Repair-PrintSystem](https://raw.githubusercontent.com/karanikn/karanik_WinMaintenance_Online/main/Screenshots/Repair-PrintSystem.png) |

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
| 1.5.1 | Clean All Restore Points | Removes all existing restore points (with confirmation countdown) |
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
| 3.1 | [Clean all Temporary files](#clean-temp-files) | Options dialog (9 steps, retention, dry run): system temp, all user profiles (Temp older than 24 h, Chrome/Edge all profiles, Teams), WU download cache, system caches, Recycle Bin, old IIS/WU logs and crash dumps, upgrade leftovers, cleanmgr with time limit, root junk; junction-safe |
| 3.2 | Clear Teams cache only | Stops Teams, clears all cache folders for all user profiles (Classic + MSIX) |
| 3.3 | Teams Full clean + reinstall | WPF GUI: selective uninstall steps + choice of Teams version to reinstall |
| 3.4 | Cleanup SoftwareDistribution | Clears SoftwareDistribution folder (preserves DataStore/history) |

### 📤 Export Tools

| ID | Script | Description |
|---|---|---|
| 4.1 | System Information | Full system report: OS, hardware, network, services, processes, event errors, remote access IDs (RustDesk / AnyDesk / TeamViewer / HopToDesk). Saved as .txt and opened in Notepad |
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
| 5.1 | Reset Print Spooler | Stops Spooler, clears spool files, restarts service (quick fix; for full diagnostics and repair use `6.10` Repair-PrintSystem) |
| 5.2 | Network Reset | Resets Winsock, TCP/IP stack, releases/renews IP, flushes DNS |
| 5.3 | Wi-Fi Reset | Full network reset + removes stale dni_dne registry key |
| 5.4 | Set Classic Right-Click Menu | Enables the classic Windows 10-style context menu in Windows 11 |
| 5.5 | Restore Default Right-Click Menu | Restores the default Windows 11 context menu |
| 5.6 | Disable Hibernation | Runs `powercfg /h off` and verifies hiberfil.sys removal |
| 5.7 | Enable Hibernation | Runs `powercfg /h on` and verifies hiberfil.sys creation |
| 5.8 | WebView2 Repair / Install | Checks, downloads and installs Microsoft Edge WebView2 Runtime |
| 5.9 | WebView2 / Outlook Auth Repair (Advanced) | Full Outlook Modern Auth / WAM / AAD BrokerPlugin repair + diagnostics export |
| 5.10 | Remove iCloud (All Users) | Removes iCloud/Apple components from all user profiles; kills related processes |
| 5.11 | [DISM / SFC Repair](#dism-sfc-repair) | WPF dialog: DISM `/CheckHealth`, `/ScanHealth`, `/RestoreHealth`, `SFC /scannow`, DISM + SFC (recommended), `/AnalyzeComponentStore`, `/StartComponentCleanup` (+ `/ResetBase`); optional `install.wim`/`.esd` repair source; live progress; SFC `[SR]` lines exported from CBS.log |

### ⭐ Extra Tools

| ID | Tool | Description |
|---|---|---|
| 6.1 | RDS Licensing Grace Period Reset | Resets the RDS Licensing Grace Period on Windows Server — can be run as many times as needed |
| 6.2 | Microsoft Activation Scripts (MAS) | Launches `irm https://get.activated.win \| iex` in a new elevated window |
| 6.3 | WinScript | Launches `irm https://winscript.cc/irm \| iex` |
| 6.4 | Chris Titus Tech's Windows Utility | Launches `irm https://christitus.com/win \| iex` |
| 6.5 | WinDiag-AI | WPF GUI for Windows diagnostics with local AI analysis (Ollama); SMART, PDF export, driver check |
| 6.6 | AD Manager | WPF GUI for common Active Directory administration tasks (requires RSAT + domain connectivity) |
| 6.7 | Network Scanner | WinForms GUI network scanner: subnet scan, MAC/vendor identification, rogue DHCP server detection |
| 6.8 | AD Domain Toolkit | Interactive console menu for AD diagnostics and confirmed repairs: replication health, DNS/DC discovery, secure channel, time sync, SYSVOL/NETLOGON checks, HTML+JSON report export (requires RSAT) |
| 6.9 | M365 User Reset | Resets Microsoft 365 authentication for a user: clears WAM accounts, Identity/Licensing registry, OneAuth/IdentityCache caches, Office credentials; optional Outlook profile reset + Autodiscover fix. Creates timestamped backup with .reg backups and full log. Requires Windows PowerShell 5.1 |
| 6.10 | Repair-PrintSystem | WPF window with live output panel (console menu on Server Core) that diagnoses, repairs, resets and restores the Windows printing subsystem: status header with Spooler state and warnings, full audit with HTML report, spooler crash analysis (faulting DLL mapped to driver/monitor), network printing checks and reversible workarounds for 0x11b / 0x709 / 0xbcb, spooler service repair, stuck jobs and spool reset, checklist removal of printers / drivers / ports / orphan monitors, full reset that keeps PDF / XPS / Fax (typed `RESET`), automatic backup before every change and restore from a backup list |

### 🏢 Office Tools

| ID | Script | Description |
|---|---|---|
| 7.1 | Office Tools | Click-to-Run update control, Quick/Online repair, SaRA Enterprise (13 scenarios), Microsoft Connectivity Check (DNS/TCP/HTTPS/Traceroute/Proxy/Firewall), Outlook/Office Diagnostics collector (HTML report, Repair Mode, Full Health Checks with SFC/DISM) |

---

<a id="dism-sfc-repair"></a>

## 🩺 DISM / SFC Repair (5.11) — Options explained

![DISM / SFC Repair dialog](https://raw.githubusercontent.com/karanikn/karanik_WinMaintenance_Online/main/Screenshots/DISM-SFC.png)

`DISM-SFC.ps1` opens a dialog where you choose **one** operation. The script then runs it, writes live progress to the launcher log and prints a summary at the end.

If nothing is clicked within **60 seconds**, the recommended option (**Repair image + system files**) starts automatically, so the tool never blocks a Run Serial sequence. Any click or key press in the dialog cancels the countdown.

### Background: what DISM and SFC actually check

Windows keeps a master copy of every system component in the **component store** (`C:\Windows\WinSxS`). The two tools work on different layers:

| Tool | Checks | Repairs from |
|---|---|---|
| **DISM** (`/Online /Cleanup-Image`) | The component store itself (the "master copies") | Windows Update, or a repair source (`install.wim` / `install.esd`) |
| **SFC** (`sfc /scannow`) | The protected system files Windows actually uses (`System32`, etc.) | The component store |

This is why the order matters: if the component store is damaged, SFC repairs files *from damaged copies* or cannot repair them at all. Microsoft therefore recommends running **DISM first, then SFC**.

### Options at a glance

| # | Option in the dialog | Command | Changes the system? | Typical duration* | Use it when |
|---|---|---|---|---|---|
| 1 | **Quick check** | `DISM /Online /Cleanup-Image /CheckHealth` | No | Seconds | You want a fast first look before deciding anything |
| 2 | **Full scan** | `DISM /Online /Cleanup-Image /ScanHealth` | No | 5–20 min | You want a reliable diagnosis without repairing |
| 3 | **Repair image** | `DISM /Online /Cleanup-Image /RestoreHealth` | Yes | 10–40 min | The scan found corruption, or Windows Update / feature installation fails |
| 4 | **Repair system files** | `sfc /scannow` | Yes | 10–30 min | Windows behaves oddly (crashes, missing features) and DISM has already been run |
| 5 | **Repair image + system files** ⭐ | DISM `/RestoreHealth`, then `sfc /scannow` | Yes | 20–60 min | **Default and recommended** for any "Windows is broken" case |
| 6 | **Analyze component store** | `DISM /Online /Cleanup-Image /AnalyzeComponentStore` | No | 1–5 min | You want to know the real WinSxS size and whether a cleanup is worth it |
| 7 | **Clean up component store** | `DISM /Online /Cleanup-Image /StartComponentCleanup` | Yes | 5–30 min | WinSxS is large and option 6 says cleanup is recommended |
| 7+ | …with **/ResetBase** | `… /StartComponentCleanup /ResetBase` | Yes, **irreversible** | 10–60 min | Maximum space recovery on a stable, fully updated system (e.g. before imaging) |

\* Durations are indicative; they depend on disk speed (SSD/HDD), CPU, the size of the component store and the internet connection.

All DISM commands are run with `/English`, so the output can be evaluated correctly on any Windows display language (e.g. Greek).

### Detailed description of each option

#### 1. Quick check — `DISM /CheckHealth`
- **What it does:** Reads whether Windows has *already recorded* the component store as corrupted (for example, by a failed update or a previous scan). It does **not** scan the files.
- **Possible results:** *No component store corruption detected* / *The component store is repairable* / *The component store cannot be repaired*.
- **Limitation:** "No corruption detected" only means that nothing has been recorded — it is not proof of a healthy system. For a real check use option 2.
- **Changes:** None.

#### 2. Full scan — `DISM /ScanHealth`
- **What it does:** Scans every component in the component store and verifies it. Details are written to `C:\Windows\Logs\CBS\CBS.log`.
- **Possible results:** Same three results as option 1, but based on an actual scan.
- **When the result is "repairable":** run option 3 or 5.
- **When the result is "cannot be repaired":** a repair install (in-place upgrade from the same Windows version ISO, keeping files and apps) is usually the next step.
- **Changes:** None.

#### 3. Repair image — `DISM /RestoreHealth`
- **What it does:** Scans the component store and **replaces damaged components** with good copies. By default the good copies are downloaded from **Windows Update**; alternatively a **repair source** can be specified (see below).
- **Progress:** The percentage can stay on the same value for several minutes. This is normal; do not stop the process.
- **Common failure:** `0x800F081F` — source files not found. Usually Windows Update cannot be reached (WSUS, proxy, Group Policy) or the source does not match the installed build. Solution: use a repair source of the same version and build.
- **Changes:** Yes, the component store is repaired. A restart may be requested (exit code `3010`).

#### 4. Repair system files — `sfc /scannow`
- **What it does:** Verifies all protected system files against the copies in the component store and **replaces any modified, damaged or missing file**.
- **Possible results (as shown by SFC):**
  - *Did not find any integrity violations* — all files are fine.
  - *Found corrupt files and successfully repaired them* — repaired; a restart is recommended.
  - *Found corrupt files but was unable to fix some of them* — run DISM `/RestoreHealth` (option 3), restart, then run SFC again.
  - *Could not perform the requested operation* — usually a pending restart; restart and run again, or run SFC from Safe Mode.
- **How the script evaluates it:** SFC's console text is translated into the Windows language, so the script does not rely on it. It reads the `[SR]` lines of **this run** from `C:\Windows\Logs\CBS\CBS.log` (always in English), counts repaired / unrepaired files and saves those lines to `DISM-SFC_<timestamp>_sfcdetails.txt` in the log folder.
- **Changes:** Yes, damaged system files are replaced.

#### 5. Repair image + system files ⭐ (recommended, default)
- **What it does:** Option 3 followed by option 4, in the order Microsoft recommends.
- **Why this is the default:** It covers both layers. Running only SFC on a damaged component store often ends with "unable to fix some of them".
- **If DISM fails:** SFC still runs (it may still repair files whose store copies are intact), and the log clearly states that DISM did not complete.
- **Changes:** Yes. A restart is recommended at the end if anything was repaired.

#### 6. Analyze component store — `DISM /AnalyzeComponentStore`
- **What it does:** Reports the **real** size of WinSxS. File Explorer over-reports it, because many files are hard links that also appear in `System32` and are counted twice.
- **Main fields in the output:**

  | Field | Meaning |
  |---|---|
  | Windows Explorer Reported Size | The size Explorer shows (hard links counted more than once) |
  | Actual Size of Component Store | The real disk usage, hard links counted once |
  | Shared with Windows | Files that are part of the running Windows — not overhead |
  | Backups and Disabled Features | Old component versions kept for rollback, and disabled features — overhead |
  | Cache and Temporary Data | Internal servicing cache — overhead |
  | Date of Last Cleanup | When the last component cleanup finished |
  | Number of Reclaimable Packages | Superseded packages a cleanup could remove |
  | Component Store Cleanup Recommended | Yes / No |

- The script reads the last field and tells you in the log whether option 7 is worth running.
- **Changes:** None.

#### 7. Clean up component store — `DISM /StartComponentCleanup`
- **What it does:** Removes **superseded** component versions (older versions replaced by newer updates). Windows does the same automatically through a scheduled task, but that task waits 30 days after an update and stops after 1 hour. This command removes them **immediately** and without a time limit.
- **Effect:** Frees disk space. Without `/ResetBase`, Windows keeps what it needs for normal servicing.
- **The script logs** the free space on the system drive before and after.
- **Changes:** Yes.

#### 7+. Option `/ResetBase` (checkbox under option 7)
- **What it does:** Removes **all** superseded versions of every component.
- ⚠️ **Irreversible:** after it completes, **updates that are already installed can no longer be uninstalled**. Updates installed later can still be uninstalled normally. The dialog asks for confirmation before it runs.
- **Use it only** on a stable system where you are sure you will not need to roll back an update (e.g. before creating a reference image or on a server with a strict disk budget).

### Repair source (optional, for options 3 and 5)

Used when Windows Update cannot provide the repair files (offline machines, WSUS, error `0x800F081F` / `0x800F0906` / `0x800F0907`).

1. Mount a Windows ISO of the **same version, edition and build** as the installed system (e.g. Windows 11 25H2, Pro).
2. Click **Browse…** and select `sources\install.wim` or `sources\install.esd` on the mounted drive.
3. The dropdown lists the editions inside the file; the one matching the installed edition is selected automatically. You can also type the index manually.
4. If the build of the selected image differs from the installed build, the dialog warns you — the repair will most likely fail.
5. **Do not use Windows Update as a repair source (`/LimitAccess`)** — when checked, DISM uses **only** your file. It is checked automatically when a file is selected.

The resulting parameter is `/Source:Wim:<path>:<index>` (or `Esd:`). A folder (e.g. a mounted image's `Windows` folder) is also accepted.

> Tip: an ISO that is older than the installed cumulative update can still fail, because the target needs newer file versions than the source contains.

### Pre-flight checks (before any operation)

| Check | Why |
|---|---|
| Pending restart (CBS / Windows Update) | SFC may refuse to run ("there is a system repair pending") and DISM results are unreliable |
| Free space on the system drive (< 5 GB) | DISM repair and cleanup need working space |
| Windows Modules Installer (TrustedInstaller) disabled | SFC and DISM repair cannot work without it |
| Another `dism.exe` already running | Two servicing operations at the same time interfere |

Warnings do not stop the run; they are written to the log so you know why a result may be unreliable.

### Recommended workflow

1. **Repair image + system files** (option 5) — covers most cases.
2. Restart Windows if anything was repaired.
3. If SFC reports files it could not repair: run option 5 again after the restart.
4. If DISM fails with `0x800F081F`: run option 3 or 5 with a **repair source** from a matching ISO.
5. If DISM reports *cannot be repaired*: a repair install (in-place upgrade) is the next step.
6. For disk space: option 6 first, then option 7 only if cleanup is recommended.

### Logs and result files

| File | Contents |
|---|---|
| `Logs\DISM-SFC_<timestamp>.log` | Full run log: pre-flight checks, commands, output, progress, summary |
| `Logs\DISM-SFC_<timestamp>_sfcdetails.txt` | `[SR]` lines from CBS.log for this SFC run (only when SFC ran) |
| `C:\Windows\Logs\DISM\dism.log` | Detailed DISM log (Windows) |
| `C:\Windows\Logs\CBS\CBS.log` | Detailed servicing / SFC log (Windows) |

Exit code of the script: `0` = completed (with or without warnings), `1` = completed with errors. The launcher allows this item up to **2 hours** (`TimeoutSec=7200`) instead of the default 1 hour.

### Command line usage (outside the launcher)

```powershell
.\DISM-SFC.ps1                                    # dialog (60 s auto-start)
.\DISM-SFC.ps1 -AutoStartSeconds 0                # dialog without auto-start
.\DISM-SFC.ps1 -Mode RestoreHealthSfc             # no dialog
.\DISM-SFC.ps1 -Mode RestoreHealth -Source "Wim:D:\sources\install.wim:6" -LimitAccess
.\DISM-SFC.ps1 -Mode ComponentCleanup -ResetBase  # irreversible, see above
```

`-Mode` values: `Menu`, `CheckHealth`, `ScanHealth`, `RestoreHealth`, `Sfc`, `RestoreHealthSfc`, `AnalyzeComponentStore`, `ComponentCleanup`.

---

<a id="clean-temp-files"></a>

## 🧹 Clean Temporary Files (3.1) — Options explained

![Clean Temporary Files dialog](https://raw.githubusercontent.com/karanikn/karanik_WinMaintenance_Online/main/Screenshots/CleanTemp.png)

`CleanTempFiles.ps1` (v3.0.1) opens a dialog where you choose **what to clean** (9 steps), how long to **keep recent files** and whether to run a **dry run**. The script then runs the selected steps in order, writes every result to the launcher log and prints a summary with the real free space before and after.

If nothing is clicked within **60 seconds**, all steps run with the default options, so the tool never blocks a Run Serial sequence. Any click or key press in the dialog cancels the countdown.

### Steps at a glance

| # | Step | What is cleaned | Safeguards |
|---|---|---|---|
| 1 | **System temp folders** | `Windows\Temp`, `Prefetch`, `LocalService\Temp`, Windows Error Reporting queues, CBS logs, `Downloaded Program Files` | Files in use are skipped |
| 2 | **User profile temp and caches** (all users) | `Temp`, IE / Edge legacy cache and cookies, WER, `CrashDumps`, thumbnail / icon cache, `Recent`, Teams classic + new Teams cache, Chrome / Edge cache (**all profiles**), Firefox `cache2` | `Temp`: only files **older than 24 h**. Teams: skipped for users who have Teams running |
| 3 | **Windows Update download cache** | `SoftwareDistribution\Download` | Skipped while Windows servicing (`TiWorker.exe`) is busy. Services stop with a 90 s limit; only services that were stopped are restarted |
| 4 | **System caches** | Font cache, DNS cache, Delivery Optimization cache, *optional* Windows Installer `$PatchCache$` | `$PatchCache$` is opt-in (see below) |
| 5 | **Recycle Bin** | All users (per-SID folders) | Only when larger than **100 MB** |
| 6 | **Old logs and crash dumps** | IIS logs (`inetpub\logs\LogFiles`), `HTTPERR` logs, Windows Update ETL logs, `dism.log`, `MEMORY.DMP`, `Minidump`, `LiveKernelReports` | Logs older than **30 days**; dumps older than **7 days**; `dism.log` only when over **100 MB** and DISM is not running |
| 7 | **Windows upgrade leftovers** | `$WINDOWS.~BT`, `$WINDOWS.~WS`, `$GetCurrent` | Only when older than **10 days**, no `Windows.old`, no setup running, no pending restart |
| 8 | **Windows Disk Cleanup** | `cleanmgr /sagerun:100` (Update Cleanup, Previous Installations, and more) | Skipped while `TiWorker.exe` is busy; time limit (**40 min**) and stall detection |
| 9 | **Root drive junk files** | `C:\*.tmp`, `*.gid`, `*.chk` (root only) | `*.bak` / `*.old` are **not** touched; `Windows.old` is only reported |

### Options

| Option in the dialog | Default | Range | Meaning |
|---|---|---|---|
| User temp files: keep newer than | 24 hours | 0–8760 | Protects files of running applications and installers (important on RDS servers). `0` = delete all |
| IIS / Windows Update logs: keep newer than | 30 days | 0–3650 | Applies to IIS, HTTPERR and Windows Update ETL logs |
| Crash dumps: keep newer than | 7 days | 0–3650 | Keeps the most recent dump for troubleshooting. `0` = delete all |
| cleanmgr time limit | 40 minutes | 5–55 | cleanmgr is terminated when the limit is reached (same behaviour as the built-in SilentCleanup task) |
| Also Windows Installer patch cache | off | — | Cleans `C:\Windows\Installer\$PatchCache$\Managed` |
| **Dry run** | off | — | Reports what *would* be deleted. Nothing is deleted, no services are stopped, cleanmgr is not started |

### Detailed description of each step

#### 1. System temp folders
- **What it does:** Deletes the contents of `C:\Windows\Temp`, `Prefetch`, `ServiceProfiles\LocalService\AppData\Local\Temp`, `ProgramData\Microsoft\Windows\WER\ReportArchive` / `ReportQueue`, the files in `Windows\Logs\CBS` and `Downloaded Program Files`.
- **Notes:** Prefetch is rebuilt automatically by Windows. The active `CBS.log` is always in use and is skipped.

#### 2. User profile temp and caches
- **Which profiles:** Every real profile under `C:\Users` (service profiles, `Default`, `Public` and junctions are excluded).
- **Temp age filter:** Only files older than the configured hours are deleted, so applications and installers that are running right now keep their files.
- **Teams:** Microsoft requires Teams to be closed before clearing its cache. The script checks which users have `Teams.exe` (classic) or `ms-teams.exe` (new Teams) running and skips their Teams cache. The new Teams folder is `AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams`.
- **Browsers:** Chrome and Edge `Cache`, `Code Cache` and `GPUCache` for every profile (`Default`, `Profile 1`, `Profile 2`, …, `Guest Profile`), Firefox `cache2` for every profile. Files locked by a running browser are skipped.

#### 3. Windows Update download cache
- **What it does:** Stops `wuauserv`, `bits` and `cryptsvc`, empties `SoftwareDistribution\Download`, then restarts **only the services it stopped**.
- **Why the TiWorker check:** While Windows servicing (`TiWorker.exe`) is installing or cleaning up updates, `wuauserv` stays in `STOP_PENDING` until servicing finishes — on Windows Server 2016 this can take an hour or more. The script measures TiWorker CPU activity for 5 s and skips the step if it is busy.
- **Stop with time limit:** Each service gets up to 90 s to stop. If `wuauserv` does not stop, the download cache is left untouched. The script never kills the service process (on Server 2016 these services share `svchost.exe` with other services).

#### 4. System caches
- **Font cache:** `FontCache` and `FontCache3.0.0.0` are stopped (30 s limit), the cache is deleted and the services are started again.
- **DNS cache:** `ipconfig /flushdns`.
- **Delivery Optimization:** `Delete-DeliveryOptimizationCache` where available (Windows 10 1709+ / Server 2019+); on Server 2016 the cache folders are cleaned directly.
- **Installer patch cache (opt-in):** Microsoft documents that the baseline cache can be deleted, but uninstalling a patch or repairing the product may then ask for the original installation media or fail. Use it only when disk space is critical.

#### 5. Recycle Bin
- Emptied for **all users** (per-SID folders under `$Recycle.Bin`) only when the total size exceeds 100 MB; `desktop.ini` is kept.

#### 6. Old logs and crash dumps
- **IIS:** `*.log` files under `C:\inetpub\logs\LogFiles` older than the retention. Sites with a custom log folder (IIS Manager → Logging) are not covered.
- **HTTPERR:** `*.log` under `C:\Windows\System32\LogFiles\HTTPERR`.
- **Windows Update:** `*.etl` traces under `C:\Windows\Logs\WindowsUpdate`.
- **dism.log:** a single file that is always recent, so age does not apply; it is deleted only when larger than 100 MB and no `dism.exe` / `DismHost.exe` is running (DISM recreates it).
- **Crash dumps:** `MEMORY.DMP` and `Minidump\*.dmp` (paths read from `HKLM\SYSTEM\CurrentControlSet\Control\CrashControl`, so custom locations are honoured) and `LiveKernelReports\*.dmp`, older than the dump retention.

#### 7. Windows upgrade leftovers
- `$WINDOWS.~BT` can contain rollback data for a recent feature update, so every folder is cleaned only when **all** of these are true: it was last changed more than 10 days ago, no Windows setup process is running, no restart is pending, and (for `$WINDOWS.~BT`) `Windows.old` does not exist.
- Protected files that cannot be deleted are reported; cleanmgr (step 8) removes the rest.

#### 8. Windows Disk Cleanup (cleanmgr)
- Sets `StateFlags0100` for the safe cleanmgr categories (including **Update Cleanup** and **Previous Installations**) and runs `cleanmgr /sagerun:100`.
- **Duration:** *Update Cleanup* is a component store cleanup performed by TiWorker. It can take 30–60 minutes, especially on Windows Server 2016.
- **Time limit:** cleanmgr is terminated after the configured minutes.
- **Stall detection:** if cleanmgr, TiWorker and DismHost use no CPU for 10 minutes, cleanmgr is considered stuck and is terminated early.
- **Progress:** every 5 minutes the log shows the elapsed time, the CPU time used in the last 5 minutes and whether TiWorker is running.
- Skipped when cleanmgr is not installed (Server Core), when another cleanmgr is already running (e.g. the SilentCleanup task), or when TiWorker is busy.

#### 9. Root drive junk files
- Deletes `*.tmp`, `*.gid` and `*.chk` in the root of the system drive only (no subfolders). `DumpStack.log.tmp` is excluded (always held by the crash-dump driver).
- `*.bak` and `*.old` are **not** deleted: on servers they can be SQL Server or user backups.
- `Windows.old` is only reported; it is removed by cleanmgr → *Previous Windows installation(s)*.

### Safety rules applied to every step

| Rule | Why |
|---|---|
| Junctions and symbolic links are **never followed**; only the link itself is removed | Windows PowerShell 5.1 `Remove-Item -Recurse` can follow a junction and delete the contents of its target |
| Files in use are skipped and counted as *locked/skipped* | No errors, no partially deleted application data |
| Services are stopped with a time limit and never killed | Avoids endless `Waiting for service ... to stop` and does not take down co-hosted services |
| Only services the script stopped are restarted | Services that were stopped before the run stay stopped |
| Dry run measures with the same code path that deletes | The reported size matches what a real run removes (except cleanmgr and DNS) |

### Logs and summary

| Item | Contents |
|---|---|
| `Logs\CleanTempFiles_<timestamp>.log` | Selected options, one result line per location (`Done` / `Would delete` / `Nothing to clean` / `Locked/skip`), warnings and summary |
| Summary: *Freed (counted)* | Bytes and files deleted by the script itself |
| Summary: *Free space C:* | Real free space before → after (includes cleanmgr; other disk activity also affects it) |

Exit code: `0` = completed (with or without warnings), `1` = completed with errors. The launcher allows this item up to **90 minutes** (`TimeoutSec=5400`).

### Command line usage (outside the launcher)

```powershell
.\CleanTempFiles.ps1                                        # dialog (60 s auto-start)
.\CleanTempFiles.ps1 -AutoStartSeconds 0                    # dialog without auto-start
.\CleanTempFiles.ps1 -DryRun -NoDialog                      # report only, no window
.\CleanTempFiles.ps1 -Steps "1,2,5" -NoDialog               # quick clean
.\CleanTempFiles.ps1 -Steps 6 -LogRetentionDays 14 -NoDialog
.\CleanTempFiles.ps1 -UserTempMinAgeHours 0 -NoDialog       # delete all user temp files
```

| Parameter | Environment variable | Default |
|---|---|---|
| `-Steps "1,2,6"` | `KARANIK_WM_CLEAN_STEPS` | all (1–9) |
| `-DryRun` | `KARANIK_WM_CLEAN_DRYRUN=1` | off |
| `-UserTempMinAgeHours` | `KARANIK_WM_CLEAN_TEMP_MIN_AGE_HOURS` | 24 |
| `-LogRetentionDays` | `KARANIK_WM_CLEAN_LOG_DAYS` | 30 |
| `-DumpRetentionDays` | `KARANIK_WM_CLEAN_DUMP_DAYS` | 7 |
| `-CleanmgrTimeoutMin` | `KARANIK_WM_CLEANMGR_TIMEOUT_MIN` | 40 |
| `-IncludePatchCache` | `KARANIK_WM_CLEAN_PATCHCACHE=1` | off |
| `-NoDialog` | `KARANIK_WM_CLEAN_NODIALOG=1` | dialog shown |
| `-AutoStartSeconds` | `KARANIK_WM_CLEAN_AUTOSTART_SEC` | 60 |

Parameters override environment variables, which override the defaults. Values chosen in the dialog override all of them.

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
| **Timeouts** | Global `$ExecTimeoutSec` (3600 s); optional per-item `TimeoutSec` in the catalog for `Remote` items (single run and Run Serial; `3.1` = 5400 s, `5.11` = 7200 s). On timeout the whole child process tree is terminated (`taskkill /T /F`) |

---

## 📝 Changelog

### v2.4 — October 2026

- **`3.1` Clean Temporary Files — rewritten (`CleanTempFiles.ps1` v3.0.1)**, see [Clean Temporary Files — Options explained](#clean-temp-files):
  - **Options dialog** (same pattern as `5.11`): choose the 9 steps, retention for user temp files / logs / crash dumps, cleanmgr time limit, Installer patch cache and **dry run**. Auto-starts with the defaults after 60 s, so Run Serial is not blocked. Also available as parameters (`-Steps`, `-DryRun`, `-NoDialog`, …) and `KARANIK_WM_CLEAN_*` environment variables for use outside the launcher
  - **Dry run** — reports per location what would be deleted and the total, without deleting anything, stopping services or starting cleanmgr
  - **New step 6 — Old logs and crash dumps:** IIS (`inetpub\logs\LogFiles`) and `HTTPERR` logs and Windows Update ETL logs older than 30 days; `dism.log` when over 100 MB and DISM is not running; `MEMORY.DMP`, `Minidump` and `LiveKernelReports` dumps older than 7 days (dump paths read from `CrashControl`)
  - **New step 7 — Windows upgrade leftovers:** `$WINDOWS.~BT`, `$WINDOWS.~WS`, `$GetCurrent`, only when older than 10 days, no `Windows.old` (rollback), no setup running and no pending restart
  - **User Temp age filter** — only files older than 24 h are deleted, so running applications and installers on RDS servers keep their files
  - **Browsers / Teams** — Chrome and Edge `Cache`, `Code Cache`, `GPUCache` for **all** profiles (previously `Default` only); new Teams (MSIX) cache added; Teams caches are skipped for users who have Teams running (Microsoft requirement)
  - **IconCache.db** cleaned for every profile (previously the current user only)
  - **Windows Update cache (step 3)** — skipped while Windows servicing (`TiWorker.exe`) is busy; services stop with a 90 s limit instead of waiting forever (`Waiting for service 'Windows Update (wuauserv)' to stop...` on Server 2016); only the services that were stopped are restarted; the service process is never killed (shared `svchost.exe`)
  - **cleanmgr (step 8)** — time limit (default 40 min) instead of waiting without limit; **stall detection** (no CPU activity of cleanmgr / TiWorker / DismHost for 10 min → terminated); progress line every 5 min with CPU used; skipped while TiWorker is busy or another cleanmgr (SilentCleanup) is running
  - **Junction-safe deletion engine** for all steps — junctions / symlinks are never followed (Windows PowerShell 5.1 `Remove-Item -Recurse` can delete the contents of a junction target); freed bytes are counted during deletion, so large folders are no longer scanned twice
  - **Summary** shows the real free space of the system drive before → after, in addition to the counted bytes
  - **Fixed:** byte counts printed to the console after every folder (`17102`, `0`, `129486620`, …)
  - **Fixed:** `Add-Content : Stream was not readable` — the log is now appended with shared read/write access and retries
  - **Fixed:** `Get-Item : Could not find item ...\IconCache.db` and the false `Deleted icon cache: 0 B` (hidden file read without `-Force`)
  - **Fixed:** Installer patch cache path — `"$PatchCache$"` inside double quotes expanded to `Installer\$\Managed`; the step is now opt-in
  - **Fixed:** Delivery Optimization cache path; `Delete-DeliveryOptimizationCache` is used where available
  - **Fixed:** font cache files stayed locked — `FontCache` / `FontCache3.0.0.0` are now stopped with a time limit before cleaning
  - **Safety — removed:** deletion of `*.bak` / `*.old` in the root of `C:` (can be SQL Server or user backups) and forced deletion of `Windows.old` with `takeown` / `icacls` / `Remove-Item` (junction risk); `Windows.old` is now reported and removed by cleanmgr *Previous Installations*
- **`6.10` Repair-PrintSystem v2.1.3 — WPF interface:** main window with status header (Spooler badge, counts, warnings), option list with descriptions, **output panel** where the messages and tables of every action appear live (Clear / Copy / Save as..., also saved to `output.txt`) and a *Last* result line; checklist dialogs for every removal (nothing pre-selected), typed confirmation for `RESET` / `YES`, workarounds dialog with current values and risk levels, backup list for restore; console menu kept as fallback (`-Console`, Server Core); log lines now appear live in the launcher log; catalog tooltip updated
- **Launcher catalog `3.1`** — new detailed tooltip; `TimeoutSec=5400` (90 min) so a long cleanmgr run is not cut by the global 3600 s limit
- **README: new section [Clean Temporary Files — Options explained](#clean-temp-files)** with the steps table, options, detailed description of each step, safety rules, logs and command line parameters
- **README: screenshots** of the Clean Temporary Files and DISM / SFC Repair dialogs (Screenshots table and top of each options section) and of the Repair-PrintSystem window

### v2.3 — October 2026

- **1 new tool added to Fix Tools:**
  - `5.11` **DISM / SFC Repair** (`DISM-SFC.ps1`) — WPF dialog to choose the operation: Quick check (`/CheckHealth`), Full scan (`/ScanHealth`), Repair image (`/RestoreHealth`), Repair system files (`SFC /scannow`), Repair image + system files (DISM then SFC, Microsoft recommended order, default), Analyze component store, Clean up component store (optional `/ResetBase` with confirmation). Optional repair source (`install.wim` / `install.esd`) with automatic image list, edition pre-selection, build mismatch warning and `/LimitAccess`. Pre-flight checks (pending restart, free space, TrustedInstaller disabled, running DISM). Live progress in the launcher log (throttled to 10% steps). DISM runs with `/English` and the SFC verdict is taken from `CBS.log` `[SR]` lines, so results are evaluated correctly on Greek Windows; the `[SR]` lines are saved as `DISM-SFC_<timestamp>_sfcdetails.txt` next to the log. SFC UTF-16 output is decoded correctly. Auto-starts the recommended option after 60 s if nothing is clicked, so it does not block Run Serial
- **README: new section [DISM / SFC Repair — Options explained](#dism-sfc-repair)** with the options table, a detailed description of each option, repair source steps, pre-flight checks, recommended workflow and log files
- **Per-item execution timeout** — catalog entries can set `TimeoutSec` to override the global 3600 s limit (`5.11` uses 7200 s); honoured for single runs and Run Serial
- **Timeout now kills the process tree** — previously only `powershell.exe` was killed and native child tools (e.g. `dism.exe`) kept running in the background

### v2.2 — October 2026

- **1 new tool added to Extra Tools:**
  - `6.10` **Repair-PrintSystem** — interactive console menu that merges `Repair-Printing.ps1` and `Reset-PrintingSystem.ps1` into one escalation ladder (diagnostics → targeted repair → full reset → restore); full audit + HTML report, spooler crash analysis, network printing checks and reversible workarounds for `0x0000011b` / `0x00000709` / `0x00000bcb`, automatic backup before every change and full restore; built-in PDF / XPS / Fax and Remote Desktop printers are protected; language independent (works on Greek Windows) (`Type=Standalone`)
- **`6.9` M365 User Reset** — expanded tooltip (modes, what is and is not changed, backup, requirements)
- **Catalog sorting fix** — IDs are now sorted as versions (`[version]`) instead of decimals (`[double]`); previously `5.10` sorted as `5.1` and appeared out of order. `1.51` Clean All Restore Points renamed to `1.5.1` to keep its position between `1.5` and `1.6`
- **Tooltips are ASCII only** — non-ASCII characters such as the em dash inside catalog strings break the launcher when Windows PowerShell 5.1 reads the file without a BOM

### v2.1 — October 2026

- **`4.1` System Information — expanded Remote Access IDs section:**
  - Replaced legacy TeamViewer/AnyDesk registry-only detection with full multi-tool detection engine
  - Detects local client IDs for **RustDesk**, **AnyDesk**, **TeamViewer**, and **HopToDesk**
  - Per-tool: reads `.toml`/`.conf` config files, queries service registry paths, spawns `--get-id` CLI (2.5 s timeout), and checks APPDATA/systemprofile locations
  - Reports: ID / Status / Source / Note per tool; explicit "not installed" when absent
  - Footer note: "Local IDs only. This is not an online-status or connectivity test."
- **2 new tools added to Extra Tools:**
  - `6.8` **AD Domain Toolkit** — interactive console menu for AD diagnostics and confirmed repairs; replication health (repadmin), DNS/DC discovery, secure channel, time sync, SYSVOL/NETLOGON, HTML+JSON export, self-test mode (`-selftest`); requires RSAT (`Type=Standalone`)
  - `6.9` **M365 User Reset** — resets Microsoft 365 authentication for a named user account; clears WAM accounts via WinRT API, Office Identity/Licensing registry keys, OneAuth/IdentityCache caches, and Office/ADAL/OneDrive credentials from Windows Credential Manager; optional Outlook profile reset + Autodiscover fix; Diagnostics mode (read-only); creates timestamped backup folder with `.reg` exports and full log; **requires Windows PowerShell 5.1** (not compatible with PS7) (`Type=Standalone`)

### v2.0 — June 2026

- **`7.1` Office Tools — major upgrade** (renamed from `Manage-OfficeUpdates.ps1` → `OfficeTool.ps1`):
  - **Single self-contained file** — Outlook/Office Diagnostics collector fully embedded (no companion script)
  - **Microsoft Connectivity Check** (`14`) — checkbox dialog to pick: DNS resolution, TCP/443, HTTPS, Traceroute (3 key hosts), Proxy config (netsh winhttp + Internet Options), Outbound Windows Firewall rules for Office/Microsoft. Runs in background, GUI stays responsive. Saves `ConnectivityCheck_timestamp.txt` report.
  - **Outlook/Office Diagnostics** (`15a/15b/15c`) — embedded HTML health report (crashes, add-ins, profiles, OST/PST, resiliency keys, WER); Repair Mode (safe Normal.dotm/dictionary reset with backup); Full Health Checks (+ SFC/DISM). All run in background.
  - **Detailed tooltips** on all 15 buttons (hover to see exact commands, registry paths, what is and is not changed)
  - **Clickable log path** in status bar — click the path to open the log file directly in Notepad
  - **Resizable log panel** — drag the bar above LOG OUTPUT to resize; color-coded output (green=OK, red=error, yellow=warn, blue=section headers)
  - **Resizable window** (`ResizeMode=CanResize`, default 860×1180px)
  - **Background task engine** (`Invoke-PSAsync`) — connectivity checks and diagnostics never freeze the GUI

### v1.9 — June 2026

- **1 new tool added to Extra Tools:**
  - `6.7` **Network Scanner** — WinForms GUI network scanner; scans subnets for live hosts, identifies MAC addresses and NIC vendors, detects unauthorized (rogue) DHCP servers (`Type=Standalone`)

### v1.8 — May 2026

- **1 new script added to Extra Tools:**
  - `6.6` **AD Manager** — WPF GUI for common Active Directory administration tasks; requires RSAT (Active Directory module) and domain connectivity (`Type=Standalone`)

### v1.7 — April 2026

- **4 new scripts added to catalog:**
  - `5.8` **WebView2 Repair / Install** — checks, downloads and installs Microsoft Edge WebView2 Runtime; kills stale msedgewebview2.exe processes before install
  - `5.9` **WebView2 / Outlook Auth Repair (Advanced)** — comprehensive repair tool for Outlook Modern Auth / WAM / AAD BrokerPlugin issues; cleans caches, exports diagnostics, tests Microsoft 365 endpoints
  - `5.10` **Remove iCloud (All Users)** — removes iCloud/Apple components (processes, AppData folders, ProgramData) from all user profiles
  - `6.5` **WinDiag-AI** — WPF GUI for Windows diagnostics with local AI analysis via Ollama; includes SMART details, PDF export, driver check, dark/light theme (`Type=Standalone`)
- **Tooltips added to all 39 catalog entries** — each tooltip describes exact steps, affected locations, and use cases; written from script source analysis

### v1.6 — March 2026

- **Office Tools group added** (`7.1` Office Tools / `Manage-OfficeUpdates.ps1`) — Click-to-Run update control (enable/disable/check/force), Quick and Online repair, Mail Control Panel, SaRA Enterprise with 13 scenarios (uninstall, Outlook scan, activation reset, Teams add-in fix, CalCheck)
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
