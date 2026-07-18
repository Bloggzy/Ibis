# Changelog

All notable Ibis changes are recorded here.

Ibis uses pre-1.0 semantic-style versioning while it is still in beta. Patch releases such as `v0.5.1` are intended for incremental project changes and small feature additions.

## v0.6.9 - 2026-07-18

- Separated Windows Event Log outputs into dedicated `EvtxECmd`, `DuckDB`, `Hayabusa`, `Takajo`, and `Chainsaw` folders beneath `HOSTNAME\EventLogs`.
- Standardised generated Event Log output files to the `HOSTNAME-Tool-...` convention, including recursive normalization of Takajo automagic output files.
- Kept Takajo backup and audit material outside its result directory so prior results can be archived reliably before a new run.
- Remove empty Takajo temporary workspaces and apply a final recursive naming pass to Takajo output files.
- Continue Takajo cleanup when Defender or another process blocks an individual output-file rename, and record the affected files as warnings.

## v0.6.8 - 2026-07-18

- Added tool install assessments that distinguish missing installs, a single versioned executable needing normalization, multiple ambiguous legacy executables, partial installs, and canonical installs with legacy files still present.
- Fixed versioned executable installs for Hayabusa and Takajo by selecting the executable from the newly staged release before publishing, rather than allowing old versioned binaries to make the post-install rename ambiguous.
- Added a confirmation-gated `Reinstall Selected` action on the Setup tools tab. It downloads the latest configured source and archives the active contents of a dedicated tool folder under `_ibis-backup` before publishing the new release.
- Preserved shared `EZTools` installations during forced reinstalls: Ibis does not clear the shared directory wholesale.

## v0.6.7 - 2026-07-17

- Added a PowerShell readiness check on the Setup tools tab that runs automatically and can be rechecked on demand.
- Checks execution-policy scopes, Group Policy enforcement, internet-origin marks on Ibis PowerShell files, and a fresh background-runspace import of `Ibis.Core.psm1` before downloads begin.
- Blocks a tool download before it starts when the background runspace cannot import Ibis, and shows the child PowerShell error instead of only the `EndInvoke` wrapper.
- Added an explicit, confirmation-gated action to remove internet-origin marks from trusted local Ibis PowerShell files. Ibis does not change execution policy or bypass Group Policy, AppLocker, or WDAC.

## v0.6.6 - 2026-06-19

- Added ordered per-user progress logging inside User artefact processing so log entries for each profile are emitted as each artefact is processed before moving to the next profile.
- Suppressed end-of-module replay of User artefact command/file hints in the background processing path to reduce confusing log duplication and out-of-context entries.
- Added broader hidden-item discovery by using `-Force` on remaining file/directory lookup paths and related `Get-Item` checks.
- Added regression coverage for hidden user artefact paths and per-user progress ordering.

## v0.6.5 - 2026-06-19

- Hardened User artefact processing so a failure in one user profile artefact does not stop later artefacts for the same user or later user profiles.
- Added per-user/per-artefact failed and skipped result records for NTUSER.dat, UsrClass.dat, Jump Lists, Recent LNKs, ShellBags, and PSReadLine processing.
- Updated processing summaries so nested skipped tool/artefact records are reported even when the parent processing module failed or completed with warnings.
- Added a regression test for an external user artefact tool launch failure on the first user profile while later profile artefacts continue processing.
- Added a regression test for a missing target artefact path in the first user profile while later artefacts and profiles continue processing.

## v0.6.4 - 2026-05-07

- Changed processing file move/rename hint logging so move records are emitted through the background progress log as each module completes, instead of being replayed in one batch at the end of the run.

## v0.6.3 - 2026-05-07

- Updated external tool process capture to drain stdout and stderr concurrently while tools run, preventing pipe-buffer hangs seen with ParseUSBs.
- Added a regression test that runs a noisy child PowerShell process and verifies Ibis captures large stdout and stderr streams without blocking.

## v0.6.2 - 2026-04-28

- Extended the end-of-run processing summary to list skipped modules/tools and their skip reasons, in addition to failed items.

## v0.6.1 - 2026-04-28

- Added an end-of-run processing summary that reports worked, failed, and skipped module counts.
- Failed modules and nested failed tool results are now listed at the end of the run so analysts know what to revisit in the logs.

## v0.6.0 - 2026-04-28

- Added a project versioning note that patch versions roll from `.9` to the next minor version rather than using `.10`.
- Added a Visual C++ Redistributable 2015+ x64 prerequisite check on the Setup tools tab.
- Added a Setup tools button that opens Microsoft's latest supported Visual C++ Redistributable page.
- Added guidance for installing the runtime with `winget install -e --id Microsoft.VCRedist.2015+.x64`.

## v0.5.9 - 2026-04-28

- Renamed the Setup tab to `Setup tools`.
- Added an `Open tools folder` button beside the tools folder browser, enabled only when the selected folder exists.
- Added admin-only controls for the Windows `LongPathsEnabled` setting.
- Shortened Defender-aware tool staging paths and extraction subfolders to reduce path length pressure during installs such as Chainsaw.
- Added a 7-Zip extraction fallback after `Expand-Archive` and .NET ZIP extraction fail.

## v0.5.8 - 2026-04-27

- Added the Info tab tagline: "A bin chicken for Windows forensic artefacts."

## v0.5.7 - 2026-04-27

- Renamed module helper folders from `Workings` to `_Working` so transparency/audit folders sort to the top and stand out as special output.
- Changed background processing command line hints so they are emitted through progress updates alongside the module that produced them, rather than being appended at the end of the run.

## v0.5.6 - 2026-04-27

- Refreshed `README.md` as a prospective user/GitHub guide for the current GUI, workflow, tools, modules, logging, configuration, and test process.
- Refreshed `AGENTS.md` as a from-scratch rebuild and development playbook aligned to the current architecture and implementation lessons.
- Updated `TODO.md` to separate completed functionality from remaining validation, hardening, packaging, and design follow-ups.

## v0.5.5 - 2026-04-27

- Normalised SrumECmd timestamped CSV outputs to the host-aware `Hostname-SrumECmd-...` filename format.
- Confirmed User Access Logs / SUM already normalises SumECmd timestamped CSV outputs to the host-aware `Hostname-SumECmd-...` filename format.

## v0.5.4 - 2026-04-27

- Refined Setup tab layout with automatic initial tool checking, clearer tool management button labels, and a larger guidance/output text area.
- Renamed Setup tool check action to `Recheck Tools` and clarified the missing-tools download button.
- Expanded Run tools log output space and updated the hostname extraction button label to identify the SYSTEM hive source.
- Adjusted About tab title styling to better match other non-hero tab titles.

## v0.5.3 - 2026-04-27

- Applied the GUI text normalisation helper consistently across the Setup page guidance/output text box.
- Normalised setup output line endings and stripped ANSI/control sequences before display for tool checks, Defender actions, downloads, and Hayabusa rule updates.
- Kept session log and output-file behaviour separate from GUI display cleanup so forensic output records are not silently rewritten.

## v0.5.2 - 2026-04-27

- Moved the About tab to the end of the tab list so core workflow tabs stay first.
- Fixed changelog display in the About tab by normalising line endings for WinForms text boxes.
- Fixed Hayabusa rule update output display by normalising line endings and stripping ANSI colour/control sequences before writing to the Setup page output box.
- Set external process stdout/stderr decoding to UTF-8 where supported, improving display of Unicode output from tools such as Hayabusa.

## v0.5.1 - 2026-04-27

- Added application version tracking in `config.json`.
- Added this changelog as the release history record.
- Added an About tab to show the current version and changelog in the GUI.

## v0.5.0 - 2026-04-26

Initial beta baseline after the major initial build of Ibis from scratch (over the Anzac long weekend 2026-04-25).

- Added a WinForms GUI for setup, tool acquisition, evidence selection, output selection, processing module selection, settings, and logs.
- Added downloader and installer support for configured DFIR tools, including GitHub latest-release handling, staged installs, backups, executable renaming, and install validation.
- Added Microsoft Defender exclusion checks, add/remove actions, and administrator-awareness for tools likely to trigger false positives.
- Added background processing for tool downloads and processing runs so the GUI stays responsive.
- Added progress reporting, pause/resume, cancel-before-next-module, completion popup notifications, and optional completion beep.
- Added session logging with ISO-style timestamps, command line hints, file operation audit records, and quick access to the logs folder/current log.
- Added source read-only boundary checks and GUI wording to remind analysts to use read-only mounts or evidence copies.
- Added host-aware output grouping, including optional blank hostname prefixes without forced `HOST` folders.
- Added System Summary processing with RegRipper.
- Added Velociraptor Results copy-out when a Results folder is present.
- Added Windows Registry hive preparation, dirty hive transaction replay attempts, prepared hive caching, and RegRipper processing.
- Added Amcache, AppCompatCache / ShimCache, Prefetch, SRUM, User artefacts, Windows Event Logs, Hayabusa, Takajo, Chainsaw, User Access Logs / SUM, BrowsingHistoryView, Forensic webhistory, ParseUSBs, and NTFS metadata modules.
- Added MFTECmd processing for `$MFT` and USN Journal `$J` discovery across mounted images and Velociraptor NTFS upload folders.
- Added DuckDB event log summaries using editable SQL query files.
- Added GUI hover hints for processing modules.
- Added Apache 2.0 licence file and GUI disclaimer/licence text.
