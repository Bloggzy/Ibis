# Agent Instructions

## Project Intent

Ibis is a Windows PowerShell DFIR orchestration tool. It prepares common forensic tooling and runs selected first-pass processing modules against Windows evidence from mounted disk images, Velociraptor/KAPE triage collections, or similar exports.

The project is a rebuild of an older single-file script. Preserve the analyst workflow knowledge, command lines, and edge cases from the old script, but keep this implementation maintainable, testable, and extendable.

Current version: `v0.6.1`.

## Build From Scratch Shape

If starting again, build toward this structure:

- `Run-Ibis.ps1`: thin launcher.
- `config.json`: app metadata, defaults, completion setting, and processing module metadata.
- `CHANGELOG.md`: release history.
- `LICENSE`: Apache 2.0 licence.
- `tools/*.json`: one tool definition per external dependency.
- `queries/eventlogs/*.sql`: editable DuckDB event log summary SQL templates.
- `modules/Ibis.Core.psm1`: testable non-GUI logic.
- `modules/Ibis.Gui.psm1`: Windows Forms GUI and runspace orchestration.
- `tests/Ibis.Core.Tests.ps1`: Pester tests for GUI-independent logic.
- `logs/`: runtime session logs, ignored by git if required.

Keep GUI/state handling separate from processing logic. If behaviour can be tested without opening the GUI, put it in `Ibis.Core.psm1`.

## Product Principles

- Keep the analyst in control.
- Prefer clear status, logs, warnings, command line hints, and manual rerun paths over silent failure.
- Treat missing artefacts as normal; return `Skipped` where possible.
- Treat missing tools or failed external commands as module `Failed`, but do not stop unrelated modules.
- Treat external DFIR tools as explicit dependencies with metadata.
- Make adding a tool or processing module straightforward.
- Do not modify source evidence directly.
- Optimise for Windows analyst VMs first.

## PowerShell Compatibility

Support Windows PowerShell 5.1 where practical. Also support PowerShell 7 during development.

Avoid PowerShell 7-only patterns unless there is a strong reason. Watch for:

- `Join-Path` against non-existent drives in tests; use `[System.IO.Path]::Combine()` for offline paths such as `E:\Evidence\Users`.
- Automatic variable names such as `$event` and `$matches`; use more specific names.
- GUI code that requires Windows-only APIs.
- Text encoding differences from external tools; normalise GUI display text at the edge.

## Evidence Handling

Assume the selected source root is read-only. Ibis can write to the output directory and working directories only.

Never modify source evidence directly. Registry hive transaction replay must operate against cached working copies. If cleaning fails or cannot be verified, record a warning and continue.

Expected source paths may include:

- `Windows\System32\config`
- `Windows\System32\winevt\Logs`
- `Windows\Prefetch`
- `Windows\appcompat\Programs`
- `Windows\System32\sru`
- `Windows\System32\LogFiles\Sum`
- `Users`

Also support Velociraptor upload layouts, including:

- `uploads\auto\C%3A`
- `uploads\ntfs\%5C%5C.%5CC%3A`

The registry hive preparation cache should be reusable for system hives and user hives such as `NTUSER.dat` and `UsrClass.dat`.

## Output Conventions

Group outputs under a hostname folder when a hostname is supplied. Avoid duplicating the hostname if the user already selected the host folder.

If hostname is blank, do not create a placeholder `HOST` folder and do not prefix filenames.

Use hostname-based output names for analyst-facing files:

```text
HOSTNAME-Tool-Or-Module-Description.ext
```

Put final outputs in the module folder or host root. Put helper files, stderr, rendered SQL, copied hives, staging artefacts, and troubleshooting summaries under `_Working` where practical.

Normalise timestamp-prefixed tool outputs after execution. Current examples include PECmd, SrumECmd, SumECmd, user artefact tools, Forensic webhistory, and Chainsaw staged outputs.

## Logging and Command Hints

Each GUI session creates a timestamped log file under `logs`.

Log:

- GUI actions and important status messages.
- Run start/finish/shutdown.
- Selected paths and module decisions.
- External command line hints with full paths and arguments.
- File operations performed by Ibis, including create, update, move/rename, remove, backup, and copy where practical.
- End-of-run summaries with worked, failed, and skipped module counts and failed module/tool names.

Command line hints should also appear in the GUI run log so users can manually retry failed tools.

## External Tool Metadata

Tool definitions should cover:

- `id`
- display name
- expected executable path
- install directory
- package type
- download URL or GitHub latest-release metadata
- manual URL
- notes
- Defender exclusion need/reason
- post-install quirks such as executable rename, nested layout, or shared EZTools folder handling

Current tool set includes:

- Eric Zimmerman: AmcacheParser, AppCompatCacheParser, EvtxECmd, JLECmd, LECmd, MFTECmd, PECmd, rla, SBECmd, SrumECmd, SumECmd.
- RegRipper.
- Hayabusa.
- Takajo.
- Chainsaw.
- DuckDB CLI.
- NirSoft BrowsingHistoryView.
- forensic-webhistory.
- parseusbs.

## Tool Acquisition Rules

Use staged installs. Do not extract directly over an existing install.

For shared install directories such as `EZTools\net9`, back up only conflicting staged items, not the whole shared folder.

Prefer staging under the final install directory for Defender-sensitive tools. Chainsaw, Hayabusa, Takajo, and rule-heavy archives can trigger Defender if extracted under `%TEMP%`. Use short staging paths such as `_s\<id>\d` and `_s\<id>\x` to reduce long-path extraction failures.

Support `.NET` ZIP extraction fallback because `Expand-Archive` can fail when the local `Microsoft.PowerShell.Archive` resources are damaged. If both PowerShell and .NET extraction fail, try 7-Zip when available.

Check for partial installs. If files exist but the expected executable is missing, report `Install Incomplete`.

Known layout quirks:

- EvtxECmd has its own `EZTools\net9\EvtxECmd` subfolder.
- Most other current EZTools sit directly under `EZTools\net9`.
- Chainsaw may require renaming a platform-specific executable to the configured expected executable.
- Takajo will not write to an existing output folder, so back up the folder first.

## Defender Handling

Some tools and rule content may trigger Defender false positives.

Important behaviours:

- Standard-user Defender exclusion checks may be incomplete or denied.
- Warn the user when not running as Administrator.
- Add/remove Defender exclusions only when running as Administrator.
- Do not treat "no exclusions found" as authoritative in a standard-user session.

## GUI Requirements

Current tab order:

1. `Info`
2. `Setup tools`
3. `Run tools`
4. `Settings`
5. `Logs`
6. `About`

GUI behaviour:

- Auto-check tools when the Setup tools tab/form opens.
- Include an Open tools folder button that is enabled only when the folder exists.
- Include admin-only Windows long path support enable/disable controls.
- Include a Visual C++ Redistributable 2015+ x64 prerequisite check and link to Microsoft's supported download page.
- Tool management buttons include `Recheck Tools`, `Download Missing Tools`, `Guidance`, and `Update Hayabusa Rules`.
- Defender controls include check/add/remove exclusions and admin-aware enablement.
- Run tools groups source, output, hostname, and processing modules.
- Evidence source label should indicate read-only intent and warn users to use read-only mounts or evidence copies.
- Processing module checkboxes have hover hints from `config.json`.
- DuckDB summaries depend on EvtxECmd.
- Takajo depends on Hayabusa.
- Disable relevant controls during background operations.
- Provide progress, pause/resume, and cancel-before-next-module controls.
- Print a processing summary at the end of each run, especially highlighting failures.
- Notify at completion with a popup and optional beep.
- Use the embedded base64 Ibis icon; no external icon file is required.
- Normalise text displayed in WinForms text boxes so CRLF and ANSI/control sequences render cleanly.

## Processing Module Pattern

Each module should:

- Resolve source paths first.
- Return `Skipped` without creating host output when the source artefact is absent.
- Create output and `_Working` folders only when it has real work to do or a failure to record.
- Resolve the required tool by ID from tool definitions.
- Return `Failed` and write summary JSON when the source exists but the tool is missing.
- Capture stderr for failed external tools where practical.
- Log full command line hints.
- Write a summary JSON with source, output, tool result, warnings, renamed outputs, and useful paths.
- Return a structured object with `ModuleId`, `Status`, `HostOutputRoot`, `OutputDirectory`, `JsonPath`, and `Message` where practical.

Avoid stopping the whole run because one module failed.

## Module Notes

- System Summary: RegRipper plugins; final text/JSON in host root; helper plugin output in `_Working`.
- Velociraptor Results: locate nearby `Results` folder and copy it under host output.
- Registry Hives: cache hives and logs, attempt dirty hive replay with `rla`, process cached copies.
- Amcache: prepare `Amcache.hve`, run AmcacheParser and RegRipper variants.
- AppCompatCache: prepare SYSTEM hive and run AppCompatCacheParser.
- Prefetch: run PECmd and rename timestamped output.
- NTFS Metadata: locate `$MFT` and `$UsnJrnl:$J` across mounted roots and Velociraptor NTFS uploads; run MFTECmd.
- SRUM: run SrumECmd with SRUDB.dat and prepared SOFTWARE hive; rename timestamped output.
- User Artefacts: process all profiles, including default/system profiles; avoid duplicate folder nesting such as `PSReadLine\PSReadLine`.
- Event Logs: run EvtxECmd.
- DuckDB Event Summaries: sub-module of EvtxECmd using editable SQL templates.
- Hayabusa: produce super-verbose JSONL timeline and support `update-rules` from Setup tools.
- Takajo: consume Hayabusa JSONL; run `automagic` and stack commands; back up output folder first.
- Chainsaw: process event logs and normalise staged output.
- UAL/SUM: run SumECmd against `Windows\System32\LogFiles\Sum`; rename timestamped output.
- Browser History: BrowsingHistoryView with offline `Users` folder.
- Forensic Webhistory: run with `--date-format iso`.
- ParseUSBs: trim trailing slashes from source path; capture stdout log.

## Versioning

Use pre-1.0 semantic-style versioning while beta:

- `v0.5.0` was the first rebuilt beta baseline.
- Patch releases such as `v0.5.9` record incremental fixes, documentation refreshes, and small additions.
- Do not create `.10` patch versions. After `v0.5.9`, roll to `v0.6.0`; after `v0.6.9`, roll to `v0.7.0`, and so on.

When changing behaviour, update:

- `config.json` version.
- `CHANGELOG.md`.
- Tests that assert the current version.

## Testing

Use Pester for testable PowerShell logic. Tests should not require real forensic tools, internet access, or real evidence images.

Run both:

```powershell
Import-Module .\modules\Ibis.Core.psm1 -Force
Import-Module .\modules\Ibis.Gui.psm1 -Force
Invoke-Pester -Path .\tests -PassThru | Select-Object TotalCount, PassedCount, FailedCount
```

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Set-Location 'C:\Tools\Ibis'; Import-Module .\modules\Ibis.Core.psm1 -Force; Import-Module .\modules\Ibis.Gui.psm1 -Force; Invoke-Pester -Path .\tests -PassThru | Select-Object TotalCount, PassedCount, FailedCount"
```

Add tests for:

- Tool definition parsing.
- Expected tool paths.
- Source path discovery.
- Skip behaviour when artefacts are absent.
- Failure behaviour when tools are absent.
- Output path and hostname handling.
- Output filename normalisation.
- Registry hive cache/replay behaviour.
- GUI-independent dependency logic.
- Text display normalisation.

## Constraints

- Do not use network access unless the user asks for implementation that requires it.
- Do not store secrets in the repo.
- Do not overwrite user work without checking first.
- Do not run external forensic tools against evidence unless the user explicitly asks.
- Do not perform destructive cleanup or source modification unless explicitly requested.
- Preserve existing user edits in the working tree.

## Workflow Notes

- Read `README.md`, `TODO.md`, `CHANGELOG.md`, and this file before implementation work.
- When using the old script as reference, extract intent, command patterns, and edge cases; do not preserve accidental complexity.
- Prefer small, reviewable changes.
- If a design choice affects analyst workflow, document the tradeoff before implementing it.

