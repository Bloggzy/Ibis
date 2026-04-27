# TODO

This roadmap reflects the current state as of `v0.5.7`. The major first-pass processing modules are implemented; remaining work is mostly validation against real evidence, hardening, packaging, and richer reporting.

## Documentation

- [x] Refresh `README.md` as a prospective GitHub/user guide.
- [x] Refresh `AGENTS.md` as the rebuild/development playbook.
- [x] Maintain `CHANGELOG.md` and `config.json` version together.
- [x] Include Apache 2.0 licence file.
- [x] Include disclaimer and licence wording in the GUI.

## Completed Foundation

- [x] Multi-file project structure.
- [x] `Run-Ibis.ps1` launcher.
- [x] `config.json` application defaults and module metadata.
- [x] Per-tool JSON definitions under `tools/`.
- [x] Editable DuckDB SQL templates under `queries/eventlogs`.
- [x] `modules/Ibis.Core.psm1` for testable core logic.
- [x] `modules/Ibis.Gui.psm1` for Windows Forms GUI.
- [x] Pester coverage for non-GUI logic.
- [x] PowerShell 7 test verification.
- [x] Windows PowerShell 5.1 test verification.
- [x] Version tracking and changelog.

## Completed GUI Features

- [x] Info tab with overview, disclaimer, licence text, and embedded Ibis logo.
- [x] Setup tab with tools folder selector.
- [x] Automatic initial tool check.
- [x] Recheck tools button.
- [x] Missing tool guidance.
- [x] Download missing tools button.
- [x] Background download/install with progress feedback.
- [x] Hayabusa rules update button and background runspace.
- [x] Defender exclusion check/add/remove controls.
- [x] Admin-aware Defender control enablement and messaging.
- [x] Run tools tab with source, output, hostname, and module controls grouped.
- [x] Evidence source read-only wording and warning.
- [x] Check source paths exist button.
- [x] Open output folder button.
- [x] Extract hostname from SYSTEM hive button.
- [x] Select all/deselect all processing modules.
- [x] Processing module hover hints.
- [x] Disable relevant controls while processing.
- [x] Background processing runspace so the GUI stays responsive.
- [x] Processing status/progress feedback.
- [x] Pause/resume and cancel-before-next-module controls.
- [x] Completion popup.
- [x] Optional completion beep.
- [x] Settings tab for completion beep.
- [x] Logs tab with open logs folder/current log buttons.
- [x] About tab with version and formatted changelog.
- [x] Text display normalisation for line endings and ANSI/control sequences.

## Completed Tool Acquisition Behaviour

- [x] Direct URL downloads.
- [x] GitHub latest release discovery for supported tools.
- [x] ZIP extraction.
- [x] `.NET` ZIP fallback when `Expand-Archive` fails.
- [x] File download installs.
- [x] Staging under install directory for Defender-sensitive tools.
- [x] Avoid extracting Defender-sensitive tools under `%TEMP%`.
- [x] Partial install detection.
- [x] Backup conflicting staged items before publish.
- [x] Avoid backing up unrelated tools in shared EZTools folder.
- [x] Post-install executable rename support.
- [x] Handle Chainsaw executable naming.
- [x] Handle EvtxECmd's own EZTools subfolder.
- [x] Defender exclusion metadata for rule-heavy tools.

## Completed Output, Logging, and Evidence Behaviour

- [x] Group export output under hostname.
- [x] Avoid duplicate `HOSTNAME\HOSTNAME` output paths.
- [x] Support blank hostname by writing directly to selected output and omitting filename prefix.
- [x] Avoid placeholder `HOST` folders.
- [x] Put final system summary text/JSON in host root.
- [x] Put supporting/working files under `_Working`.
- [x] Use hostname-based output names for implemented modules.
- [x] Rename timestamp-prefixed outputs from tools such as PECmd, SrumECmd, SumECmd, and user artefact tools.
- [x] Use `_Working` folder names for special helper/transparency output folders.
- [x] Emit processing command line hints through progress updates near the related module output.
- [x] Treat absent artefact sources as skipped, not fatal.
- [x] Treat missing tools as failed module results with summary JSON.
- [x] Capture stderr for failed external tools where practical.
- [x] Create a date-stamped session log file.
- [x] Record command line hints for external tool invocations.
- [x] Record Ibis file audit events for create/update/move/remove operations where practical.
- [x] Log GUI shutdown.
- [x] Preserve tools/source/output paths and completion beep setting in `config.json`.

## Completed Processing Modules

- [x] System summary from registry hives using RegRipper.
- [x] Velociraptor `Results` copy-out.
- [x] Windows Registry hives with RegRipper.
- [x] Dirty registry hive checks.
- [x] Registry transaction replay using `rla.exe` against cached working copies.
- [x] Registry hive preparation cache for reuse across modules.
- [x] Amcache.
- [x] AppCompatCache/ShimCache.
- [x] Prefetch.
- [x] NTFS metadata with MFTECmd for `$MFT` and USN Journal `$J`.
- [x] SRUM.
- [x] User artefacts.
- [x] Windows Event Logs with EvtxECmd.
- [x] DuckDB summaries over EvtxECmd output.
- [x] Hayabusa.
- [x] Takajo.
- [x] Chainsaw.
- [x] User Access Logs/SUM.
- [x] Browser history with BrowsingHistoryView.
- [x] Browser history with forensic-webhistory.
- [x] ParseUSBs.

## Validation With Real Evidence

- [ ] Confirm EvtxECmd output naming and useful default exports with representative evidence.
- [ ] Confirm DuckDB event log summaries against large EvtxECmd CSVs.
- [ ] Confirm Hayabusa command switches and rule update workflow against representative event logs.
- [ ] Confirm Takajo outputs with representative Hayabusa JSONL.
- [ ] Confirm Chainsaw rule/mapping paths across current and future release layouts.
- [ ] Confirm SumECmd output names with real SUM databases.
- [ ] Confirm SrumECmd output names and content with real SRUM databases.
- [ ] Confirm BrowsingHistoryView behaviour on mounted/offline browser artefacts for all target browsers.
- [ ] Confirm forensic-webhistory output naming and content on Velociraptor/KAPE style collections.
- [ ] Confirm ParseUSBs output naming and useful CSV inventory with real evidence.
- [ ] Confirm MFTECmd `$MFT` and `$J` discovery across mounted images and Velociraptor NTFS upload folders.
- [ ] Confirm user artefact coverage for default/system profiles and regular user profiles.

## Hardening

- [ ] Add schema/config validation for `config.json` and `tools/*.json`.
- [ ] Add clearer handling for malformed tool definitions.
- [ ] Add external tool version capture where possible.
- [ ] Add online version checks for GitHub-backed tools.
- [ ] Add upgrade workflow for already-installed tools.
- [ ] Add better handling for direct-download tools that do not expose versions.
- [ ] Add retry/cancel controls for long downloads.
- [ ] Add per-module elapsed time.
- [ ] Add richer progress parsing from long-running external tools.
- [ ] Add a run manifest covering selected modules, commands, tool versions, output paths, warnings, and errors.
- [ ] Add command rerun hints to module summary JSON files where missing.
- [ ] Add log export or transcript bundle.
- [ ] Add cleanup policy for old `_ibis-backup` and staging folders.
- [ ] Review where file audit logging could be made more comprehensive without excessive noise.

## Packaging and Operations

- [ ] Add setup/bootstrap instructions for a fresh analyst VM.
- [ ] Add recommended PowerShell execution policy guidance.
- [ ] Add install/update instructions for PowerShell 7 while preserving PS5.1 compatibility.
- [ ] Add recommended Defender exclusion guidance for common tool folders.
- [ ] Add release packaging instructions.
- [ ] Add CI for Pester tests.
- [ ] Add sample/synthetic evidence fixtures for safer regression testing.
- [ ] Add screenshots or a short user walkthrough for the README once the UI settles.

## Design Decisions To Revisit

- [ ] Should processing modules remain hard-coded in the GUI switch, or move to a more declarative dispatch table?
- [ ] Should each processing module be split into its own `.psm1` as `Ibis.Core.psm1` grows?
- [ ] Should downloads support explicit "force reinstall" and "upgrade only if newer" modes?
- [ ] Should there be a "prepare evidence" phase that only cleans/caches hives before all modules?
- [ ] Should module output folder names avoid spaces everywhere, or preserve familiar names from the old script?
- [ ] Should skipped modules write JSON summaries, or is no output preferable when evidence is absent?
- [ ] Should the GUI expose a per-module advanced settings panel?

