# Changelog

All notable Ibis changes are recorded here.

Ibis uses pre-1.0 semantic-style versioning while it is still in beta. Patch releases such as `v0.5.1` are intended for incremental project changes and small feature additions.

## v0.7.5 - 2026-08-20

- Added Boobook as a processing module. It parses USB artefacts into a self-contained analyst report, plus CSV and JSONL of every device, event, and correlation. Ibis runs it with `-in-place` so the results land in `USB\Boobook` rather than a run directory Ibis did not name, and moves any earlier result into `USB\_Boobook-Output-Backups` first, because Boobook refuses to write into a folder that already holds anything.
- ParseUSBs output moved from `USB\` into `USB\ParseUSBs\`, so the two USB tools sit beside each other the way the event log and browser history tools already do. Existing output folders are not migrated; a re-run writes to the new location.
- The `Run tools` module list is now grouped by the artefact each module reads: system and registry, Windows Event Logs, browser history, USB artefacts. Groups come from a new `group` field in `config.json`, which decides both the reading order and the run order. The columns are chosen to keep every module on screen without scrolling, which gave the run log a taller box.
- Module names lost the words the group heading now carries, so `Windows Event Logs (EvtxECmd)` is `EvtxECmd` and `Browser history (BrowsingHistoryView)` is `BrowsingHistoryView`. The run log uses these names.
- The run no longer filters modules against a hand-written list of ids, which silently dropped any module nobody remembered to add to it. It filters on the `implemented` status already recorded in `config.json`, and a test asserts every implemented module has a dispatch case.

- Added a Buy Me a Coffee sponsor link. It appears as the GitHub Sponsor button through `.github/FUNDING.yml`, as a `Support` section in the README, and as a link on the `About` tab. Supporting Ibis is optional and buys no priority, no support obligation, and no influence over what Ibis does.
- Repaired the manual verification command in the README, where a tab character sat in place of the backslash in `.\tests\manual`, so the command as printed could not run.

## v0.7.4 - 2026-08-20

- A processing module that throws now writes a failure summary to `_Ibis-Failures\HOSTNAME-<module>-Failure.json` under the output root, recording the error message, exception type, script stack trace, and the source, tools, and output paths used. Previously a throwing module left nothing on disk, so the output folder gave no sign it had ever run. The run log names the file.
- The failure recorder never raises an error of its own, so one failing module cannot take down the rest of the run.
- Removed 364 lines of unreachable code from the `Run tools` button handler. A bare `return` sat above a second, full module dispatch chain that had not run since processing moved into a background runspace. Editing it had no effect, which made it a trap. Processing modules are now dispatched from one place only, and tests enforce that.
- Added support for Hayabusa 4.0.0, which merged `csv-timeline` and `json-timeline` into `dfir-timeline` and renamed `--ISO-8601` to `--iso-8601`. Both of the old commands now exit 2, so Hayabusa stopped working for anyone who installed or reinstalled it. Ibis asks the installed executable which commands it has and builds the right arguments, so Hayabusa 3.x and 4.x both work. The detected version, command, and detection method are recorded in the Hayabusa summary JSON.
- Hayabusa is now always run with `-C`. Without it Hayabusa refuses to overwrite an existing timeline but still exits 0 and writes nothing, so re-running a host into the same output folder silently kept the previous timeline and reported success. This affected Hayabusa 3.x as well.
- Reorganised the `Setup tools` tab so it reads in the order the work has to happen. The left column now holds the numbered machine-preparation steps: PowerShell readiness, Defender exclusions, Windows long paths, and runtime prerequisites. Tool management moved to the right column as step 5. Long path controls moved out of the tool management group into their own group.
- Every setup step now shows a status symbol in its title: a green tick when it is ready, an amber warning when it needs attention, and a red mark when it blocks work. Tool management shows the warning until the Defender exclusions are confirmed.
- `Download Missing Tools` now checks the Defender exclusions first and warns before downloading if they are missing, cannot be confirmed from a standard-user session, or could not be checked. The analyst can still continue, and the choice is recorded in the log. Downloading before adding exclusions let Defender quarantine Hayabusa and Chainsaw rule content mid-extraction, which made the tools look broken rather than blocked.
- Tool management now shows the Defender exclusion state next to the download button, so it is visible before the button is pressed.

## v0.7.3 - 2026-08-20

- Cancelling a processing run now stops the external tool that is currently running instead of waiting for it to finish. Ibis stops the process and its children, records the reason in the captured stderr, and reports `Cancelled` on the tool result.
- Added an optional `-TimeoutSeconds` to external tool capture, with process-tree termination when it is reached. The default remains no timeout, because forensic parsers can legitimately run for hours.
- Closing the Ibis window while a processing run, tool download, or Hayabusa rule update is active now asks for confirmation, then requests cancellation and stops and disposes the background runspaces instead of ending them mid-write.
- Removed a second, unreachable definition of `Resolve-IbisComparablePath`. Two copies with different behaviour existed and the later one silently won.
- Defender exclusion checks no longer treat two blank paths as a match.
- Fixed the Takajo `stack-users` failure. Takajo maps `-s` to `skipProgressBar` for every stack command except `stack-users`, where it means `sourceUsers`. Ibis was therefore leaving the progress bar enabled for that one command, which crashes with `Error: unhandled exception: The handle is invalid.` because Ibis redirects the console. Ibis now passes `--skipProgressBar` by its long name for `automagic` and every stack command, and asks for target users rather than source users.
- Takajo now writes each operation's stderr to `EventLogs\Takajo\_Working\HOSTNAME-Takajo-<mode>.stderr.txt` when the tool says anything, and includes a short excerpt of the error in the summary JSON message.
- Takajo failure and warning messages now name the summary JSON path and the operations that failed, instead of saying only "See summary JSON for details".
- Clarified the Takajo warning text: it now says Takajo writes no CSV when nothing matched, rather than implying the output is missing.
- Added regression coverage for timeout termination, analyst cancellation, cancellation checks that call back into the caller scope, single-definition path comparison, tool error excerpting, and the Takajo progress bar flag.

## v0.7.2 - 2026-08-20

- Removed the ParseUSBs explicit hive pass. ParseUSBs 1.8 crashes in that mode with `NameError: name 'userfolders' is not defined`, even with the syntax from its own help text, so it could never produce output. The module now runs one staged volume-mode pass, which covers the same registry hives and adds USB event logs and user LNK files.
- Accept Velociraptor percent-encoded USB event log names such as `Microsoft-Windows-Partition%254Diagnostic.evtx`, and stage them under the canonical Windows name so ParseUSBs can find them. The selected name and discovery method are recorded in the USB summary JSON.
- Report "no USB device connections found" as a completed result with an explicit message, instead of a warning about missing output files.
- Simplified ParseUSBs output names to `HOSTNAME-ParseUSBs-<file>` now that there is a single pass.

## v0.7.1 - 2026-08-20

- Stage ParseUSBs registry configuration files and transaction logs beneath the module working directory before the tool runs, ensuring any transaction replay remains isolated from source evidence.
- Clear the read-only attribute on Ibis working copies of evidence files so registry transaction replay by `rla` and ParseUSBs can write to the cached copy. Source evidence keeps its original attributes and is never written to.
- Added `tests\manual\Verify-ParseUsbContainment.ps1`, a manual whole-module check that proves ParseUSBs processing leaves a read-only evidence tree unchanged and is only ever given staged paths.

## v0.7.0 - 2026-07-22

- Expanded the DuckDB user-session activity timeline with workstation lock/unlock, screen saver, shutdown/restart, operating-system boundary, and clean/unexpected shutdown evidence.
- Corrected Window Station reconnect/disconnect and NTLM credential-validation labels so they do not overstate RDP or domain-controller semantics.
- Normalised Event ID, channel, username, logon type, Security logon ID, terminal session ID, related session, session name, and disconnect-reason fields across EvtxECmd event maps.
- Retained event provenance and raw evidence fields in the DuckDB result, including computer, provider, record identifiers, source file, user SID, and payload.
- Publish successful DuckDB results from a per-run staged CSV so a failed rerun preserves, and is clearly distinguished from, any previous analyst-facing output.
- Added regression coverage for the expanded event set, correlation fields, provenance columns, and current version.
- Expanded USN Journal `$J` discovery to identify bounded `$Extend` candidates extracted as `$UsnJrnl_$J`, `$UsnJrnl-$J`, standalone `$J`, or flattened `$UsnJrnl`, in addition to the existing mounted-volume and encoded Velociraptor forms.
- Prefer canonical collection names over alternates and record the selected USN Journal discovery method and confidence in the NTFS metadata summary JSON.
- Added regression coverage for renamed USN Journal candidates and canonical-name selection.

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
- Refreshed `AGENTS.md` as a development playbook aligned to the current architecture and implementation lessons.
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
