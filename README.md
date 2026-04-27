# Ibis

Ibis is a Windows PowerShell DFIR helper for analyst workstations. It can download and prepare a set of common forensic tools, then run selected first-pass parsing modules against a Windows evidence source such as a mounted disk image, Velociraptor collection, KAPE collection, or similar triage export.

Ibis is intended to speed up the early stages of analysis by producing a consistent, organised output folder. It does not replace analyst review, validation, or deeper case-specific examination.

Licensed under the Apache License, Version 2.0. Provided AS IS, without warranties or conditions of any kind. Use at your own risk.

## Status

Current version: `v0.5.7`

Ibis is pre-1.0 beta software. The current version and default settings are stored in `config.json`, and notable changes are recorded in `CHANGELOG.md`.

## Requirements

- Windows analyst VM or workstation.
- Windows PowerShell 5.1 or PowerShell 7.
- Internet access for tool downloads and GitHub latest-release checks, unless tools are installed manually.
- Administrator rights if you want Ibis to add or remove Microsoft Defender exclusions.
- A read-only mounted image or a working copy of a triage collection is strongly recommended.

Ibis treats the evidence source as read-only and performs hive transaction replay against cached working copies. However, Ibis cannot control every behaviour of every external forensic tool. For best forensic hygiene, mount images read-only or process a copy of a triage pack.

## Quick Start

From the project folder:

```powershell
.\Run-Ibis.ps1
```

Typical workflow:

1. Open the `Setup` tab and confirm the tools folder, normally `C:\DFIR\Tools`.
2. Let Ibis check installed tools automatically, then use `Download Missing Tools` or `Guidance` as needed.
3. If running as Administrator, check/add Defender exclusions for tools that commonly trigger false positives.
4. Open the `Run tools` tab.
5. Select the evidence source root, for example a mounted Windows volume root or a triage folder containing `Windows` and `Users`.
6. Select an output folder.
7. Extract the hostname from the offline SYSTEM hive or enter a hostname manually.
8. Select processing modules, or use Select All/Deselect All.
9. Click `Run selected modules`.
10. Review the output folder, session log, and per-module JSON summaries.

## GUI Tabs

- `Info`: overview, disclaimer, licence note, and Ibis logo.
- `Setup`: tools folder, tool checks, downloads, guidance, Hayabusa rule updates, and Defender exclusions.
- `Run tools`: source selection, output selection, hostname, module selection, progress, pause/resume, and cancel-before-next-module.
- `Settings`: completion notification settings, including the optional audible beep.
- `Logs`: current session log location with buttons to open the log file or logs folder.
- `About`: current version and changelog.

The GUI keeps tool downloads, Hayabusa rule updates, and processing runs in background runspaces so the form remains responsive. Processing progress is reported through status text, progress bar updates, and a scrolling run log.

## Evidence Sources

Ibis expects the source root to resemble the root of a Windows volume or a preserved triage tree. Common paths include:

- `Windows\System32\config`
- `Windows\System32\winevt\Logs`
- `Windows\Prefetch`
- `Windows\appcompat\Programs`
- `Windows\System32\sru`
- `Windows\System32\LogFiles\Sum`
- `Users`

Velociraptor collections are also supported where artefacts are stored below uploaded paths such as `uploads\auto\C%3A` and NTFS special files below `uploads\ntfs\%5C%5C.%5CC%3A`.

Missing artefacts are normal. Ibis usually records those modules as `Skipped` rather than treating the whole run as failed.

## Output Layout

Outputs are grouped by hostname when a hostname is supplied:

```text
C:\Export\HOSTNAME\<Module>\...
```

If the selected output path is already the host folder, Ibis avoids creating duplicate paths such as `HOSTNAME\HOSTNAME`. If the hostname field is blank, Ibis writes directly under the selected output folder and omits the hostname prefix from output filenames.

Most analyst-facing output files use:

```text
HOSTNAME-Tool-Or-Module-Description.ext
```

Examples:

- `HOSTNAME-RR-System-Summary.txt`
- `HOSTNAME-RR-System-Summary.json`
- `HOSTNAME-EZ-Amcache.csv`
- `HOSTNAME-SrumECmd-AppResourceUseInfo_Output.csv`
- `HOSTNAME-Hayabusa-EventLogs-SuperVerbose.jsonl`
- `HOSTNAME-Takajo-stack-logons.csv`
- `HOSTNAME-BrowsingHistoryView-All-Users.csv`
- `HOSTNAME-ForensicWebHistory-results.csv`
- `HOSTNAME-MFTECmd-MFT-Output.csv`
- `HOSTNAME-ParseUSBs-Log.txt`

Intermediate files, rendered SQL, stderr captures, copied hives, and helper outputs are stored under `_Working` folders where practical. The underscore keeps these transparency/audit folders at the top of normal file listings.

## Processing Modules

All processing modules are enabled by default in `config.json`. They can be selected or deselected in the `Run tools` tab.

### System Summary

Uses RegRipper against offline registry hives to extract core host details such as hostname, Windows version/build, install date, last shutdown, timezone, and IP/domain information.

### Velociraptor Results Copy-Out

Looks for a nearby Velociraptor `Results` folder and copies it into the case output when present.

### Windows Registry Hives

Copies Windows system hives and transaction logs into a working cache, checks whether hives are dirty, attempts transaction replay with Eric Zimmerman's `rla.exe`, then processes cached copies with RegRipper. Source evidence is not modified.

### Amcache

Prepares `Windows\appcompat\Programs\Amcache.hve`, then runs AmcacheParser and RegRipper outputs.

### AppCompatCache / ShimCache

Prepares the offline `SYSTEM` hive and runs AppCompatCacheParser.

### Prefetch

Runs PECmd against `Windows\Prefetch`. Timestamp-prefixed PECmd outputs are renamed into the hostname-based format.

### NTFS Metadata

Uses MFTECmd to process `$MFT` and, where found, USN Journal `$J`. Ibis searches mounted image roots and Velociraptor NTFS upload locations.

### SRUM

Runs SrumECmd against `Windows\System32\sru\SRUDB.dat` with a prepared `SOFTWARE` hive. Timestamp-prefixed SrumECmd outputs are renamed into the hostname-based format.

### User Artefacts

Processes all discovered user profile folders, including default/system profiles. Modules include RegRipper user hive output, Jump Lists, Recent LNKs, ShellBags, PSReadLine history, Run keys, and UserAssist.

### Windows Event Logs

Runs EvtxECmd against `Windows\System32\winevt\Logs`.

### DuckDB Event Log Summaries

Optional sub-module of Windows Event Logs. It consumes EvtxECmd CSV output and runs editable SQL templates from `queries\eventlogs` to produce summary CSVs such as logon and outbound RDP pivots.

### Hayabusa

Runs Hayabusa against Windows event logs and produces a super-verbose JSONL timeline. The `Setup` tab can also run Hayabusa's rule update workflow.

### Takajo

Consumes Hayabusa JSONL output. Takajo is disabled unless Hayabusa is selected. Ibis runs `automagic` plus explicit stack commands and backs up any existing Takajo output folder first because Takajo will not write into an existing output directory.

### Chainsaw

Runs Chainsaw against Windows event logs using bundled rule content. Outputs are staged and normalised to hostname-based event log files.

### User Access Logs / SUM

Runs SumECmd against `Windows\System32\LogFiles\Sum`. Timestamp-prefixed SumECmd outputs are renamed into the hostname-based format. These artefacts are normally found on Windows Server systems.

### Browser History

Runs NirSoft BrowsingHistoryView against offline user browser history artefacts.

### Forensic Webhistory

Runs `forensic-webhistory scan -d <source> -o <output> --date-format iso` to provide an additional browser history parser with ISO date output.

### ParseUSBs

Runs parseusbs to extract USB artefact information. The module is labelled `ParseUSBs` in the GUI.

## Tool Management

Tool definitions live in `tools/*.json`. They define the tool ID, expected executable path, install directory, download source, manual URL, package type, and any known quirks.

Current tool set includes:

- Eric Zimmerman tools: AmcacheParser, AppCompatCacheParser, EvtxECmd, JLECmd, LECmd, MFTECmd, PECmd, rla, SBECmd, SrumECmd, SumECmd.
- RegRipper.
- Hayabusa.
- Takajo.
- Chainsaw.
- DuckDB CLI.
- NirSoft BrowsingHistoryView.
- forensic-webhistory.
- parseusbs.

Ibis supports direct downloads and GitHub latest-release downloads where configured. Installs are staged before publishing to avoid extracting over partial installs. Defender-sensitive tools are staged under their install directory rather than `%TEMP%` where possible.

## Defender Exclusions

Tools and rule sets such as Chainsaw, Hayabusa, and Takajo may trigger Defender false positives. The `Setup` tab can check, add, and remove recommended folder exclusions based on tool metadata.

Standard-user Defender checks may be incomplete. Administrator rights are required to add or remove exclusions.

## Logging

Each GUI session creates a timestamped log under `logs` using an ISO-style filename such as:

```text
2026-04-27T07-55-00Z.log
```

Logs include:

- GUI actions and status messages.
- Processing progress.
- Command line hints for external tools.
- File creation, move, rename, update, and removal audit events where Ibis performs them.
- Shutdown entry when the GUI closes.

The `Logs` tab can open the current log file or the logs directory.

## Configuration

`config.json` stores:

- Application name and version.
- Default tools, source, and output paths.
- Default hostname placeholder.
- Completion beep setting.
- Processing module list, labels, hover hints, default enabled state, and implementation status.

Ibis updates the tools/source/output paths and completion beep setting when changed in the GUI so the next launch resumes from the previous choices.

## Development and Tests

Pester tests cover the non-GUI core logic and should pass in both PowerShell 7 and Windows PowerShell 5.1.

PowerShell 7:

```powershell
Import-Module .\modules\Ibis.Core.psm1 -Force
Import-Module .\modules\Ibis.Gui.psm1 -Force
Invoke-Pester -Path .\tests -PassThru | Select-Object TotalCount, PassedCount, FailedCount
```

Windows PowerShell 5.1:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Set-Location 'C:\Tools\Ibis'; Import-Module .\modules\Ibis.Core.psm1 -Force; Import-Module .\modules\Ibis.Gui.psm1 -Force; Invoke-Pester -Path .\tests -PassThru | Select-Object TotalCount, PassedCount, FailedCount"
```

As of `v0.5.7`, both test runs pass with `115` tests.

