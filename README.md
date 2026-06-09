# WingetUpgradeAll.ps1

## Overview
`WingetUpgradeAll.ps1` is a PowerShell script that simplifies software management using the Windows Package Manager (`winget`). It provides four operations:

1. **List** — Display all upgradeable programs available via `winget`.
2. **Save** — Save the IDs of upgradeable programs to a file, excluding specified exceptions.
3. **Upgrade** — Upgrade programs from a previously saved list, excluding exceptions.
4. **Upgrade-all** — List live and upgrade everything in one step, excluding exceptions.

When the [`Microsoft.WinGet.Client`](https://www.powershellgallery.com/packages/Microsoft.WinGet.Client) PowerShell module is installed, the script uses it to enumerate upgradeable packages — this is robust and locale-independent. Otherwise it falls back to parsing `winget upgrade` text output. The upgrades themselves always run through the `winget` CLI, and failures are detected via the process exit code (not exceptions).

## Requirements
- **Windows** with the **Windows Package Manager (`winget`)** installed (ships with *App Installer* from the Microsoft Store).
- **PowerShell 5.1** or later (PowerShell 7+ recommended).
- *(Optional, recommended)* the `Microsoft.WinGet.Client` module for robust package enumeration:
  ```powershell
  Install-Module Microsoft.WinGet.Client -Scope CurrentUser
  ```

## Usage
Open a PowerShell terminal and run the script with a command and any optional parameters.

### Syntax
```powershell
.\WingetUpgradeAll.ps1 <list|save|upgrade|upgrade-all> [options]
```

The command is a positional value — type it directly (e.g. `.\WingetUpgradeAll.ps1 list`), not the literal word `command`.

> **Note:** If your execution policy blocks scripts, run with
> `powershell -ExecutionPolicy Bypass -File .\WingetUpgradeAll.ps1 list`.

### Commands
| Command | Description |
| --- | --- |
| `list` | Lists all upgradeable programs. |
| `save` | Saves upgradeable program IDs to a file, excluding exceptions. |
| `upgrade` | Upgrades programs from the saved list, excluding exceptions. |
| `upgrade-all` | Lists live and upgrades everything, excluding exceptions. |

### Options
| Option | Description | Default |
| --- | --- | --- |
| `-listPath <path>` | File to save (`save`) or read (`upgrade`) program IDs. | `$PSScriptRoot\UpgradeablePrograms.txt` |
| `-exceptionsPath <path>` | File of program IDs to exclude. | `$PSScriptRoot\PermanentUpgradeExceptions.txt` |
| `-logPath <path>` | Error log file. | `$PSScriptRoot\WingetUpgradeErrors.log` |
| `-Source <winget\|msstore>` | Restrict to a single winget source. | *(all sources)* |
| `-Interactive` | Run upgrades interactively instead of silently. | *(silent)* |
| `-WhatIf` | Show what would be upgraded without performing any upgrades. | *(off)* |

## The save → upgrade workflow
`save` and `upgrade` are two halves of one workflow:

1. Run `save` to capture the current upgradeable IDs to a file (review/edit it if you like).
2. Run `upgrade` later to apply exactly that list.

If you just want to upgrade everything available right now without an intermediate file, use `upgrade-all`.

## Exceptions / list file format
Both the exceptions file and the saved list file are plain text, **one program ID per line**. Blank lines and lines beginning with `#` are ignored, so you can annotate entries:

```text
# Pinned — manage manually
Mozilla.Firefox

# Breaks our plugin; skip for now
Notepad++.Notepad++
```

## Examples

### 1. List upgradeable programs
```powershell
.\WingetUpgradeAll.ps1 list
```

### 2. Save upgradeable programs (excluding exceptions)
```powershell
.\WingetUpgradeAll.ps1 save -listPath "C:\Path\To\UpgradeablePrograms.txt" -exceptionsPath "C:\Path\To\Exceptions.txt"
```

### 3. Upgrade from a saved list
```powershell
.\WingetUpgradeAll.ps1 upgrade -listPath "C:\Path\To\UpgradeablePrograms.txt"
```

### 4. Upgrade everything now, winget source only, dry run
```powershell
.\WingetUpgradeAll.ps1 upgrade-all -Source winget -WhatIf
```

## Notes
- Upgrades run silently by default with package and source agreements accepted; use `-Interactive` to override.
- The script logs errors to `WingetUpgradeErrors.log` in the script directory (configurable via `-logPath`).
- It exits with code `1` if any upgrade fails, winget is missing, or the saved list can't be found — useful for Task Scheduler / CI.
- Use meaningful paths for `-listPath` and `-exceptionsPath` to avoid overwriting important files.

## Changelog
See [CHANGELOG.md](CHANGELOG.md) for the version history.

## License
This script is provided as-is, without warranty of any kind. Use it at your own risk.
