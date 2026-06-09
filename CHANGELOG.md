# Changelog

All notable changes to this project are documented here. This project adheres to
[Semantic Versioning](https://semver.org/).

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
