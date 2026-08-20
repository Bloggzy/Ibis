# Ibis

Ibis is a Windows PowerShell DFIR helper for analyst workstations: a bin chicken for Windows forensic artefacts. It can download and prepare a set of common forensic tools, then run selected first-pass parsing modules against a Windows evidence source such as a mounted disk image, Velociraptor collection, KAPE collection, or similar triage export.

Ibis is intended to speed up the early stages of analysis by producing a consistent, organised output folder. It does not replace analyst review, validation, or deeper case-specific examination.

Licensed under the Apache License, Version 2.0. Provided AS IS, without warranties or conditions of any kind. Use at your own risk.

## Status

Current version: `v0.7.4`

Ibis is pre-1.0 beta software. The current version and default settings are stored in `config.json`, and notable changes are recorded in `CHANGELOG.md`.

## Requirements

- Windows analyst VM or workstation.
- Windows PowerShell 5.1 or PowerShell 7.
- Microsoft Visual C++ Redistributable 2015+ x64 for several native DFIR tools.
- Internet access for tool downloads and GitHub latest-release checks, unless tools are installed manually.
- Administrator rights if you want Ibis to add or remove Microsoft Defender exclusions.
- A read-only mounted image or a working copy of a triage collection is strongly recommended.

Ibis treats the evidence source as read-only and performs hive transaction replay against cached working copies. However, Ibis cannot control every behaviour of every external forensic tool. For best forensic hygiene, mount images read-only or process a copy of a triage pack.

## Getting the Files

Download the release ZIP, then extract it before you run anything.

Windows marks files that came from the internet with a "mark of the web" flag. If you extract the ZIP with the built-in Windows extractor (`Extract All` in File Explorer), that mark is copied onto every extracted file. PowerShell then treats the scripts as untrusted and shows security warnings, or refuses to run them, depending on your execution policy.

Pick one of the two options below.

### Option 1: Extract with 7-Zip (recommended)

7-Zip does not copy the mark of the web onto the extracted files, so there is nothing to clean up afterwards.

The simplest way is from File Explorer: right-click the ZIP, then choose `7-Zip` and `Extract to "Ibis\"`. On Windows 11 the `7-Zip` entry may sit under `Show more options`.

To do the same from a command line, note that the 7-Zip installer does not add `7z.exe` to `PATH`, so call it by its full path:

```powershell
& "$env:ProgramFiles\7-Zip\7z.exe" x Ibis.zip -oC:\Tools\Ibis
```

If 7-Zip was installed as 32-bit on a 64-bit host, it is under `${env:ProgramFiles(x86)}\7-Zip\7z.exe` instead.

Any archive tool that preserves this behaviour works. Only the built-in Windows extractor is a problem.

### Option 2: Unblock the PowerShell files

If you already extracted with the Windows extractor, unblock the three PowerShell files. From the project folder:

```powershell
Get-ChildItem -Path .\Run-Ibis.ps1, .\modules\Ibis.Core.psm1, .\modules\Ibis.Gui.psm1 | Unblock-File
```

The three files are:

- `Run-Ibis.ps1`
- `modules\Ibis.Core.psm1`
- `modules\Ibis.Gui.psm1`

You can also right-click each file, choose `Properties`, tick `Unblock`, and click `OK`.

Unblocking the ZIP before extracting has the same effect:

```powershell
Unblock-File -Path .\Ibis.zip
```

If the GUI does start, the `Setup tools` tab reports any remaining internet-origin marks on Ibis files and offers to remove them after confirmation. Ibis does not change execution policy and does not bypass Group Policy, AppLocker, or WDAC.

## Quick Start

Extract the release as described in `Getting the Files` first, then from the project folder:

```powershell
.\Run-Ibis.ps1
```

Typical workflow:

1. Open the `Setup tools` tab and confirm the tools folder, normally `C:\DFIR\Tools`.
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
- `Setup tools`: tools folder, then a numbered left column of machine preparation steps (PowerShell readiness, Defender exclusions, Windows long paths, runtime prerequisites) and a right column for tool management (tool checks, downloads, install assessments, selected-tool reinstall, guidance, Hayabusa rule updates).

Each setup step carries a status symbol in its title: a green tick when it is ready, an amber warning when it needs attention, and a red mark when it blocks work. The symbols are built from Unicode code points rather than written into the source, because Windows PowerShell reads a script without a byte order mark as ANSI.
- `Run tools`: source selection, output selection, hostname, module selection, progress, pause/resume, and cancel.
- `Settings`: completion notification settings, including the optional audible beep.
- `Logs`: current session log location with buttons to open the log file or logs folder.
- `About`: current version and changelog.

The GUI keeps tool downloads, Hayabusa rule updates, and processing runs in background runspaces so the form remains responsive. Processing progress is reported through status text, progress bar updates, and a scrolling run log.

Cancelling a run stops the external tool that is running at the time, then ends the run. Partial output from a stopped tool may remain in the output folder, and the stop is recorded in the log and in the tool stderr capture. Closing the Ibis window while a run, download, or rule update is still going asks for confirmation first, because closing ends the background work immediately.

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

Ibis uses hidden-aware file and directory discovery where it enumerates artefacts, because hidden files and folders can be relevant to forensic analysis.

## Output Layout

Outputs are grouped by hostname when a hostname is supplied:

```text
C:\Export\HOSTNAME\<Module>\...
```

Windows Event Log processors keep their final results in separate folders:

```text
C:\Export\HOSTNAME\EventLogs\EvtxECmd\
C:\Export\HOSTNAME\EventLogs\DuckDB\
C:\Export\HOSTNAME\EventLogs\Hayabusa\
C:\Export\HOSTNAME\EventLogs\Takajo\
C:\Export\HOSTNAME\EventLogs\Chainsaw\
```

Browser-history processors likewise retain separate results beneath one parent folder:

```text
C:\Export\HOSTNAME\WebHistory\BrowsingHistoryView\
C:\Export\HOSTNAME\WebHistory\ForensicWebHistory\
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
- `EventLogs\Hayabusa\HOSTNAME-Hayabusa-EventLogs-SuperVerbose.jsonl`
- `EventLogs\Takajo\HOSTNAME-Takajo-stack-logons.csv`
- `WebHistory\BrowsingHistoryView\HOSTNAME-BrowsingHistoryView-All-Users.csv`
- `WebHistory\ForensicWebHistory\HOSTNAME-ForensicWebHistory-results.csv`
- `HOSTNAME-MFTECmd-MFT-Output.csv`
- `USB\HOSTNAME-ParseUSBs-usb-info.csv`
- `USB\_Working\HOSTNAME-ParseUSBs-Log.txt`

Intermediate files, rendered SQL, stderr captures, copied hives, and helper outputs are stored under `_Working` folders where practical. Empty Takajo workspaces are removed after processing.

A module that fails outright writes no summary of its own, so Ibis records one for it in `_Ibis-Failures\HOSTNAME-<module>-Failure.json` under the output root. It holds the module id and name, the error message, the exception type, the script stack trace, and the source, tools, and output paths used. The run log names the file. Every selected module therefore accounts for itself in the output folder, whether it succeeded or not.

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

Uses MFTECmd to process `$MFT` and, where found, USN Journal `$J`. Ibis searches mounted image roots and Velociraptor NTFS upload locations, including common extracted names such as `$UsnJrnl_$J`, `$UsnJrnl-$J`, and standalone `$J` files within `$Extend`.

### SRUM

Runs SrumECmd against `Windows\System32\sru\SRUDB.dat` with a prepared `SOFTWARE` hive. Timestamp-prefixed SrumECmd outputs are renamed into the hostname-based format.

### User Artefacts

Processes all discovered user profile folders, including default/system profiles. Modules include RegRipper user hive output, Jump Lists, Recent LNKs, ShellBags, PSReadLine history, Run keys, and UserAssist. User artefact processing isolates failures per user and artefact so a failure in one parser does not stop later artefacts or later user profiles. Progress/log entries are emitted as each user artefact step runs, keeping each profile's entries together before the next profile starts.

### Windows Event Logs

Runs EvtxECmd against `Windows\System32\winevt\Logs`.

### DuckDB Event Log Summaries

Optional sub-module of Windows Event Logs. It consumes EvtxECmd CSV output and runs editable SQL templates from `queries\eventlogs` to produce summary CSVs such as user-session activity and outbound RDP pivots. The user-session timeline includes logon/logoff, lock/unlock, screen saver, Window Station/RDP session, privilege, authentication, and system-boundary evidence while retaining source provenance and correlation identifiers.

### Hayabusa

Runs Hayabusa against Windows event logs and produces a super-verbose JSONL timeline, plus a CSV timeline. The `Setup tools` tab can also run Hayabusa's rule update workflow.

Hayabusa 4.0.0 replaced the `csv-timeline` and `json-timeline` commands with a single `dfir-timeline` command selected by `-t csv|json|jsonl`, and renamed `--ISO-8601` to `--iso-8601`. Ibis supports both eras. It asks the installed executable which commands it has, then builds the matching arguments, so Hayabusa 3.x and 4.x both work with no configuration. The detected version, the command used, and how it was detected are recorded in the Hayabusa summary JSON.

Ibis always passes `-C`. Without it Hayabusa refuses to overwrite an existing timeline but still exits with code 0 and writes nothing, so re-running a host into the same output folder would silently keep the previous timeline.

### Takajo

Consumes Hayabusa JSONL output. Takajo is disabled unless Hayabusa is selected. Ibis runs `automagic` plus explicit stack commands and backs up any existing Takajo output folder first because Takajo will not write into an existing output directory. Its summary is kept with the Takajo results; an empty temporary workspace is removed after processing.

Ibis passes `--skipProgressBar` by its long option name. Takajo maps the short `-s` to `skipProgressBar` for most commands but to `sourceUsers` for `stack-users`, and a progress bar crashes Takajo when Ibis redirects the console.

A stack command that finds nothing writes no CSV. Ibis records that as a warning, not a failure. When Takajo writes anything to stderr, Ibis saves it to `EventLogs\Takajo\_Working\HOSTNAME-Takajo-<mode>.stderr.txt` and puts a short excerpt in the summary JSON.

### Chainsaw

Runs Chainsaw against Windows event logs using bundled rule content. Outputs are staged and normalised to hostname-based event log files.

### User Access Logs / SUM

Runs SumECmd against `Windows\System32\LogFiles\Sum`. Timestamp-prefixed SumECmd outputs are renamed into the hostname-based format. These artefacts are normally found on Windows Server systems.

### Browser History

Runs NirSoft BrowsingHistoryView against offline user browser history artefacts. Its results are written beneath `WebHistory\BrowsingHistoryView`.

### Forensic Webhistory

Runs `forensic-webhistory scan -d <source> -o <output> --date-format iso` to provide an additional browser history parser with ISO date output. Its results are written beneath `WebHistory\ForensicWebHistory`.

### ParseUSBs

Runs parseusbs to extract USB artefact information without writing to the evidence source. Ibis creates a selective, volume-shaped copy in `USB\_Working\ParseUSBs-Evidence-Staging`: system hives and transaction logs, each available `NTUSER.dat` and its logs, user LNK files, and the two USB-relevant event logs when present. Read-only attributes are cleared on the copies so the tool can replay transactions there instead of failing. parseusbs then runs once, in volume mode, against that staging directory. Any hive replay, permission change, or `.restored` file it produces stays in the staging area and the source evidence is unchanged.

Velociraptor percent-encodes the `%` in event log channel names, so `Microsoft-Windows-Partition%4Diagnostic.evtx` is collected as `Microsoft-Windows-Partition%254Diagnostic.evtx`. Ibis accepts either form and stages it under the canonical Windows name, because parseusbs looks for that exact name. The selected name and discovery method are recorded in the USB summary JSON.

Ibis does not use parseusbs explicit hive mode (`-s` / `-w` / `-u`). Version 1.8 crashes in that mode with `NameError: name 'userfolders' is not defined`, even when given the syntax from its own help text, so it can produce no output. Volume mode covers the same registry hives and adds the USB event logs and user LNK files.

When parseusbs finds no USB device connections it writes no CSV. Ibis reports that as a completed result with an explicit "no USB device connections were observed" message, not as a failure. The module is labelled `ParseUSBs` in the GUI.

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

Ibis supports direct downloads and GitHub latest-release downloads where configured. Installs are staged before publishing to avoid extracting over partial installs. Defender-sensitive tools are staged under their install directory rather than `%TEMP%` where possible, using short `_s\<id>\d` and `_s\<id>\x` staging paths to reduce path-length pressure.

ZIP extraction tries PowerShell `Expand-Archive`, then a .NET fallback, then 7-Zip if `7z.exe` or `7za.exe` is available. The `Setup tools` tab also includes admin-only controls for the Windows `LongPathsEnabled` registry setting, which may require a restart before every process observes the change.

Some tools require the Microsoft Visual C++ Redistributable 2015+ x64 runtime. Ibis checks for it on the `Setup tools` tab and links to Microsoft's supported download page. On a fresh analyst VM, it can often be installed with:

```powershell
winget install -e --id Microsoft.VCRedist.2015+.x64
```

Before downloads, Ibis checks that a fresh background PowerShell runspace can import its core module. The Setup tools tab also identifies execution-policy warnings and internet-origin marks on Ibis files. When a trusted local checkout is marked as downloaded from the internet, Ibis can remove that mark after confirmation. Ibis will not change execution policy or attempt to bypass Group Policy or application-control policy.

## Defender Exclusions

Tools and rule sets such as Chainsaw, Hayabusa, and Takajo may trigger Defender false positives. The `Setup tools` tab can check, add, and remove recommended folder exclusions based on tool metadata.

Standard-user Defender checks may be incomplete. Administrator rights are required to add or remove exclusions.

## Logging

Each GUI session creates a timestamped log under `logs` using an ISO-style filename such as:

```text
2026-04-27T07-55-00Z.log
```

Logs include:

- GUI actions and status messages.
- Processing progress.
- End-of-run processing summaries showing worked, failed, and skipped module counts plus failed and skipped item names/reasons, including nested per-user artefact failures/skips.
- Command line hints for external tools.
- Concurrent stdout/stderr capture for external tools so noisy tools do not hang on full output pipes.
- Move/rename hints emitted during the relevant module rather than replayed at run completion.
- User artefact progress emitted in profile order so per-user artefact entries remain grouped in the log.
- File creation, move, rename, update, and removal audit events where Ibis performs them.
- Shutdown entry when the GUI closes.

The `Logs` tab can open the current log file or the logs directory.

## Configuration

`config.json` stores:

- Application name and version.
- Shared, generic default tools, source, and output paths.
- Default hostname placeholder.
- Completion beep setting.
- Processing module list, labels, hover hints, default enabled state, and implementation status.

Ibis keeps the shared `config.json` defaults generic. When changed in the GUI, tools/source/output paths and the completion-beep setting are saved to the local, Git-ignored `config.local.json` so the next launch resumes from the previous choices without placing case paths in the repository.

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

As of `v0.7.4`, both test runs pass with `195` tests.

`tests\manual` holds slower whole-module checks that are run by hand rather than as part of the Pester suite. `Verify-ParseUsbContainment.ps1` proves the ParseUSBs module cannot alter source evidence: it builds a synthetic read-only evidence tree, substitutes a stand-in parser that records its command line, runs the module, and compares a SHA-256 and last-write snapshot of the source tree taken before and after. It needs Windows PowerShell 5.1 and no real tools, evidence, or network access.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\manual\Verify-ParseUsbContainment.ps1
```

## Support

Ibis is free and Apache-2.0 licensed. If it saves you time on a case, you can [buy me a coffee](https://buymeacoffee.com/bloggz). There is also a Sponsor button at the top of the repository page.

Support is optional and buys no priority, no support obligation, and no influence over what Ibis does.
