# Changelog

All notable Ibis changes are recorded here.

Ibis uses pre-1.0 semantic-style versioning while it is still in beta. Patch releases such as `v0.5.1` are intended for incremental project changes and small feature additions.

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

Initial beta baseline after the major rebuild of Ibis from scratch (over the Anzac long weekend 2026-04-25).

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
