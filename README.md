# WingetUpgradeAll.ps1

## Overview
`WingetUpgradeAll.ps1` is a PowerShell script designed to simplify software management using the Windows Package Manager (`winget`). It provides three main functionalities:

1. **List**: Display all upgradeable programs available via `winget`.
2. **Save**: Save a list of upgradeable programs to a file, excluding specified exceptions.
3. **Upgrade**: Upgrade programs from a saved list, excluding specified exceptions.

## Usage
To use the script, open a PowerShell terminal and run the script with the appropriate parameters.

### Syntax
```powershell
.\WingetUpgradeAll.ps1 command <list|save|upgrade> -listPath <path> -exceptionsPath <path>
```

### Parameters
- **`command`** (required): Specifies the operation to perform. Valid values are:
  - `list`: Lists all upgradeable programs.
  - `save`: Saves upgradeable programs to a file, excluding exceptions.
  - `upgrade`: Upgrades programs from the saved list, excluding exceptions.

- **`-listPath`** (optional): Path to save or read the list of upgradeable programs. Defaults to:
  ```
  $PSScriptRoot\UpgradeablePrograms.txt
  ```

- **`-exceptionsPath`** (optional): Path to a file containing program IDs to exclude from upgrades. Defaults to:
  ```
  $PSScriptRoot\PermanentUpgradeExceptions.txt
  ```

## Examples

### 1. List Upgradeable Programs
```powershell
.\WingetUpgradeAll.ps1 list
```
This command lists all programs that can be upgraded using `winget`.

### 2. Save Upgradeable Programs to a File
```powershell
.\WingetUpgradeAll.ps1 save -listPath "C:\Path\To\UpgradeablePrograms.txt" -exceptionsPath "C:\Path\To\Exceptions.txt"
```
This command saves the list of upgradeable programs to the specified file, excluding any programs listed in the exceptions file.

### 3. Upgrade Programs from a Saved List
```powershell
.\WingetUpgradeAll.ps1 upgrade -listPath "C:\Path\To\UpgradeablePrograms.txt" -exceptionsPath "C:\Path\To\Exceptions.txt"
```
This command upgrades programs listed in the specified file, excluding any programs listed in the exceptions file.

## Notes
- Ensure that `winget` is installed and configured on your system.
- The script logs errors to a file named `WingetUpgradeErrors.log` in the script's directory.
- Use meaningful paths for `-listPath` and `-exceptionsPath` to avoid overwriting important files.

## License
This script is provided as-is, without warranty of any kind. Use it at your own risk.
