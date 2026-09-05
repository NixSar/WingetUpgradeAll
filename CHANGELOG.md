# Changelog

All notable changes to this project are documented here. This project adheres to
[Semantic Versioning](https://semver.org/).

## [2.1.0] - 2026-09-05

### Added
- Scope-aware upgrades. Each package's install scope is determined via `winget list --scope user|machine`. Every package is upgraded under the right token regardless of how the script was started: from an elevated terminal, machine-scope packages run in-process and user-scope/unknown packages run in one de-elevated child (launched via the desktop shell, no prompt); from a normal terminal, user-scope packages run in-process and machine-scope packages run in one elevated child (one UAC prompt). Child results are merged into the summary and error log.
- `list` shows a `Scope` column; `-WhatIf` shows the non-elevated / elevated partition.
- Internal `_batch` command with `-resultPath` / `-pidPath` parameters used by the child runs.

### Changed
- `ConvertFrom-WingetUpgradeText` generalised to `ConvertFrom-WingetTableText` (also parses `winget list`); `Get-WingetUpgradeText` generalised to `Get-WingetText -Arguments`.

### Why
- Running per-user installers (Electron/Squirrel apps) under an elevated token makes them write Administrators-owned keys with wrong permissions into the user hive. On 2026-05-31 this left `HKCU\Software\Classes\.webp` unwritable by the user, which made Claude Desktop's MSIX auto-update fail registration (`0x80073CF6`) and loop, force-closing the app repeatedly.

## [2.0.5] - 2026-06-10

### Fixed
- Packages with undeterminable installed versions are no longer listed or saved; behavior matches plain `winget upgrade` (as in 1.x). Removed the `--include-unknown` flag introduced in 2.0 and applied the equivalent filter to the module path.

## [2.0.4] - 2026-06-10

### Fixed
- Separator blank line after each upgrade was swallowed when winget left the cursor mid-line; the line is now closed before the separator is written.

## [2.0.3] - 2026-06-09

### Changed
- Blank line after each upgrade for readable output.

## [2.0.2] - 2026-06-09

### Changed
- During an upgrade, winget now draws its progress directly to the console instead of being piped through PowerShell. This fixes garbled progress output (mojibake) and repeated progress-bar lines, and renders the progress bar in place.

### Fixed
- A package that is already current (winget exit code `0x8A15002B`) is now reported as "Already up to date" and counted as skipped, instead of being logged as a failure.

## [2.0.1] - 2026-06-09

### Fixed
- `upgrade` and `upgrade-all` crashed with "The '++' operator works only on numbers" on the first program. winget's console output was leaking into the upgrade function's return value; it is now routed to the host so the status is a single value.

## [2.0] - 2026-06-09

### Added
- `upgrade-all` command: enumerate upgradeable programs and upgrade them in one step.
- `-WhatIf` dry-run mode that reports what would be upgraded without changing anything.
- `-Source` filter to restrict operations to a single winget source (`winget` or `msstore`).
- `-Interactive` switch; upgrades otherwise run silently with package and source agreements accepted.
- `-logPath` option to set the error log location.
- Comment (`#`) and blank-line support in the exceptions and saved-list files.
- Run summary (succeeded / failed / skipped) and a non-zero exit code when any upgrade fails, winget is missing, or the saved list is not found.
- winget presence check with a clear message when it is not installed.
- Preferred enumeration via the `Microsoft.WinGet.Client` PowerShell module when available, with an automatic fallback to text parsing.
- Comment-based help (`Get-Help .\WingetUpgradeAll.ps1`).

### Changed
- `upgrade` no longer enumerates available updates first; it reads the saved list directly.
- Saved list is written as UTF-8.
- Temporary capture file is written to `%TEMP%` (per-process) instead of the current directory.

### Fixed
- Upgrade failures are now detected via the winget exit code instead of a `catch` that never triggered.
- Error log now records the correct program ID instead of the error object.
- Text parsing no longer crashes on short lines and tolerates blank lines, repeated headers, and non-English summaries.
- Saved list no longer attempts to upgrade empty IDs.
- Internal process object no longer leaks into parsed output.
- Corrected the script name and command syntax shown in usage output and documentation.

## [1.0] - 2025-04-22

### Added
- Initial release with `list`, `save`, and `upgrade` commands.
- Configurable list and exceptions file paths.
- Error logging to `WingetUpgradeErrors.log`.
