<#
.SYNOPSIS
    Lists, saves, and bulk-upgrades programs via the Windows Package Manager (winget),
    with support for a permanent exceptions list.

.DESCRIPTION
    WingetUpgradeAll.ps1 wraps winget to make unattended bulk upgrades easy.

    Commands:
      list         Show all upgradeable programs.
      save         Save upgradeable program IDs to a file, excluding exceptions.
      upgrade      Upgrade every program listed in a previously saved file.
      upgrade-all  List live and upgrade everything in one step, excluding exceptions.

    When the Microsoft.WinGet.Client PowerShell module is installed it is used to
    enumerate upgradeable packages (robust, locale-independent). Otherwise the script
    falls back to parsing `winget upgrade` text output. Upgrades themselves always run
    through the winget CLI so failures are detected via the process exit code.

.PARAMETER command
    Operation to perform: list, save, upgrade, or upgrade-all.

.PARAMETER listPath
    Path used to save (save) or read (upgrade) the list of program IDs.
    Defaults to <script dir>\UpgradeablePrograms.txt.

.PARAMETER exceptionsPath
    Path to a file of program IDs to exclude. One ID per line; blank lines and
    lines starting with '#' are ignored. Defaults to
    <script dir>\PermanentUpgradeExceptions.txt.

.PARAMETER logPath
    Path to the error log. Defaults to <script dir>\WingetUpgradeErrors.log.

.PARAMETER Source
    Restrict to a single winget source: 'winget' or 'msstore'.

.PARAMETER Interactive
    Run upgrades interactively instead of silently.

.PARAMETER WhatIf
    Show what would be upgraded without performing any upgrades.

.EXAMPLE
    .\WingetUpgradeAll.ps1 list

.EXAMPLE
    .\WingetUpgradeAll.ps1 save -exceptionsPath .\Exceptions.txt

.EXAMPLE
    .\WingetUpgradeAll.ps1 upgrade

.EXAMPLE
    .\WingetUpgradeAll.ps1 upgrade-all -Source winget -WhatIf

.NOTES
    Version: 2.0.5
#>
#Requires -Version 5.1

param (
    [Parameter(Position = 0)]
    [ValidateSet("list", "save", "upgrade", "upgrade-all")]
    [string] $command,

    [string] $listPath       = "$PSScriptRoot\UpgradeablePrograms.txt",
    [string] $exceptionsPath  = "$PSScriptRoot\PermanentUpgradeExceptions.txt",
    [string] $logPath         = "$PSScriptRoot\WingetUpgradeErrors.log",

    [ValidateSet("winget", "msstore")]
    [string] $Source,

    [switch] $Interactive,
    [switch] $WhatIf
)

$script:ExitCode  = 0
$script:UseModule = $null   # resolved lazily on first use

# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #

function Write-Log {
    param ([string] $Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $logPath -Value "$timestamp  $Message"
}

function Test-WingetInstalled {
    return [bool] (Get-Command winget -ErrorAction SilentlyContinue)
}

function Test-WinGetModule {
    # Cache the result so we only probe / import once per run.
    if ($null -ne $script:UseModule) {
        return $script:UseModule
    }

    if (Get-Module -Name Microsoft.WinGet.Client) {
        $script:UseModule = $true
        return $true
    }

    if (Get-Module -ListAvailable -Name Microsoft.WinGet.Client) {
        Import-Module Microsoft.WinGet.Client -ErrorAction SilentlyContinue
        $script:UseModule = [bool] (Get-Module -Name Microsoft.WinGet.Client)
        return $script:UseModule
    }

    $script:UseModule = $false
    return $false
}

function Get-Field {
    param (
        [string] $Line,
        [int]    $Start,
        [int]    $End
    )

    if ($Start -lt 0 -or $Start -ge $Line.Length) { return '' }

    $len = if ($End -lt 0 -or $End -gt $Line.Length) { $Line.Length - $Start } else { $End - $Start }
    if ($len -le 0) { return '' }

    return $Line.Substring($Start, $len).Trim()
}

function ConvertFrom-WingetUpgradeText {
    param ([string[]] $OutputLines)

    $results     = [System.Collections.Generic.List[object]]::new()
    $headerIndex = -1
    $cols        = $null

    # Locate the header line and capture column start positions.
    for ($i = 0; $i -lt $OutputLines.Count; $i++) {
        if ($OutputLines[$i] -match 'Name\s+Id\s+Version\s+Available\s+Source') {
            $line        = $OutputLines[$i]
            $headerIndex = $i
            $cols = [ordered]@{
                Name      = $line.IndexOf('Name')
                Id        = $line.IndexOf('Id')
                Version   = $line.IndexOf('Version')
                Available = $line.IndexOf('Available')
                Source    = $line.IndexOf('Source')
            }
            break
        }
    }

    if ($headerIndex -lt 0) {
        return $results
    }

    # Process rows after the header (header + separator line).
    for ($i = $headerIndex + 2; $i -lt $OutputLines.Count; $i++) {
        $line = $OutputLines[$i]

        if ([string]::IsNullOrWhiteSpace($line))                      { continue }
        if ($line -match '^\d+\s+upgrade')                            { break }    # "N upgrades available."
        if ($line -match 'Name\s+Id\s+Version\s+Available\s+Source')  { continue } # repeated header (2nd table)
        if ($line -match '^[\s\-─—]+$')                     { continue } # separator rule

        $id = Get-Field -Line $line -Start $cols.Id -End $cols.Version

        # Real winget IDs never contain whitespace; this filters prose / banners.
        if ([string]::IsNullOrWhiteSpace($id) -or $id -match '\s') { continue }

        $results.Add([pscustomobject]@{
            Name      = Get-Field -Line $line -Start $cols.Name      -End $cols.Id
            Id        = $id
            Version   = Get-Field -Line $line -Start $cols.Version   -End $cols.Available
            Available = Get-Field -Line $line -Start $cols.Available -End $cols.Source
            Source    = Get-Field -Line $line -Start $cols.Source    -End -1
        })
    }

    return $results
}

function Get-WingetUpgradeText {
    # Capture `winget upgrade` output to a private temp file (avoids CWD clutter
    # and name collisions between concurrent runs).
    $tempFile = Join-Path $env:TEMP "winget_upgrade_$PID.txt"
    try {
        $null = Start-Process -FilePath 'winget' `
            -ArgumentList 'upgrade' `
            -NoNewWindow -Wait `
            -RedirectStandardOutput $tempFile
        return Get-Content -Path $tempFile -Encoding UTF8
    }
    finally {
        if (Test-Path $tempFile) {
            Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-UpgradeableItems {
    param ([string] $Source)

    if (Test-WinGetModule) {
        # Skip packages whose installed version winget cannot determine; they
        # cannot be meaningfully upgraded (matches `winget upgrade` default).
        $packages = Get-WinGetPackage | Where-Object {
            $_.IsUpdateAvailable -and $_.InstalledVersion -and $_.InstalledVersion -ne 'Unknown'
        }
        if ($Source) {
            $packages = $packages | Where-Object { $_.Source -eq $Source }
        }
        return $packages | ForEach-Object {
            [pscustomobject]@{
                Name      = $_.Name
                Id        = $_.Id
                Version   = $_.InstalledVersion
                Available = ($_.AvailableVersions | Select-Object -First 1)
                Source    = $_.Source
            }
        }
    }

    # Fallback: parse CLI text output.
    $items = ConvertFrom-WingetUpgradeText -OutputLines (Get-WingetUpgradeText)
    if ($Source) {
        $items = $items | Where-Object { $_.Source -eq $Source }
    }
    return $items
}

function Get-CleanLines {
    # Read a file, trim each line, drop blanks and '#' comments.
    param ([string] $Path)

    if (-not (Test-Path $Path)) { return @() }

    return Get-Content -Path $Path -ErrorAction SilentlyContinue |
        ForEach-Object { $_.Trim() } |
        Where-Object   { $_ -and -not $_.StartsWith('#') }
}

function Invoke-Upgrade {
    param (
        [string] $Id,
        [switch] $Interactive,
        [switch] $DryRun
    )

    if ($DryRun) {
        Write-Host "  [WhatIf] Would upgrade: $Id"
        return 'Skipped'
    }

    $wingetArgs = @(
        'upgrade', '--id', $Id, '--exact',
        '--accept-package-agreements', '--accept-source-agreements'
    )
    $wingetArgs += if ($Interactive) { '--interactive' } else { '--silent' }

    Write-Host "  Upgrading: $Id"
    # Let winget draw directly to the console (correct encoding, in-place
    # progress). Capturing the process object keeps it out of the function's
    # return value; -Wait makes ExitCode available.
    $proc = Start-Process -FilePath 'winget' -ArgumentList $wingetArgs `
        -NoNewWindow -Wait -PassThru
    $code = $proc.ExitCode

    # 0x8A15002B: package is already current. Not a failure.
    if ($code -eq -1978335189) {
        Write-Host "  Already up to date: $Id"
        return 'Skipped'
    }

    # winget is a native exe: it signals failure via exit code, not exceptions.
    if ($code -ne 0) {
        $message = "Failed to upgrade '$Id' (winget exit code $code)."
        Write-Host "  $message" -ForegroundColor Red
        Write-Log $message
        return 'Failed'
    }

    return 'Succeeded'
}

function Invoke-UpgradeSet {
    param ([string[]] $Ids)

    if (-not $Ids -or @($Ids).Count -eq 0) {
        Write-Host "Nothing to upgrade."
        return
    }

    $tally = @{ Succeeded = 0; Failed = 0; Skipped = 0 }

    foreach ($id in $Ids) {
        $status = Invoke-Upgrade -Id $id -Interactive:$Interactive -DryRun:$WhatIf
        $tally[$status]++
        # winget's progress rendering can leave the cursor mid-line; close
        # that line first so the separator below is actually blank.
        try {
            if ($Host.UI.RawUI.CursorPosition.X -gt 0) { Write-Host "" }
        } catch {}
        Write-Host ""
    }

    Write-Host ("Summary: {0} succeeded, {1} failed, {2} skipped." -f `
        $tally.Succeeded, $tally.Failed, $tally.Skipped)

    if ($tally.Failed -gt 0) {
        $script:ExitCode = 1
    }
}

function Show-Usage {
    Write-Host "Usage:"
    Write-Host "  .\WingetUpgradeAll.ps1 <list|save|upgrade|upgrade-all> [options]"
    Write-Host ""
    Write-Host "Commands:"
    Write-Host "  list         Lists all upgradeable programs."
    Write-Host "  save         Saves upgradeable program IDs to a file, excluding exceptions."
    Write-Host "  upgrade      Upgrades programs from the saved list, excluding exceptions."
    Write-Host "  upgrade-all  Lists live and upgrades everything, excluding exceptions."
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  -listPath <path>        List file to save/read. Default: $listPath"
    Write-Host "  -exceptionsPath <path>  Program IDs to exclude.   Default: $exceptionsPath"
    Write-Host "  -logPath <path>         Error log.                Default: $logPath"
    Write-Host "  -Source <winget|msstore>  Restrict to one source."
    Write-Host "  -Interactive            Upgrade interactively instead of silently."
    Write-Host "  -WhatIf                 Show what would be upgraded without doing it."
}

# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

if (-not $command) {
    Show-Usage
    exit 0
}

if (-not (Test-WingetInstalled)) {
    Write-Host "winget was not found. Install 'App Installer' from the Microsoft Store and retry." -ForegroundColor Red
    Write-Log "winget command not found."
    exit 1
}

try {
    $exceptions = Get-CleanLines -Path $exceptionsPath

    switch ($command) {

        "list" {
            $items = Get-UpgradeableItems -Source $Source
            if (-not $items -or @($items).Count -eq 0) {
                Write-Host "No upgrades available."
                break
            }
            $items | Format-Table Name, Id, Version, Available, Source -AutoSize
            Write-Host ("{0} upgrade(s) available." -f @($items).Count)
        }

        "save" {
            $items  = Get-UpgradeableItems -Source $Source
            $toSave = @($items | Where-Object { $_.Id -notin $exceptions })
            if ($toSave.Count -eq 0) {
                Write-Host "Nothing to save (no upgrades, or all excluded)."
                break
            }
            $toSave | ForEach-Object {
                Write-Host "Saving: $($_.Name) [$($_.Id)]"
                $_.Id
            } | Set-Content -Path $listPath -Encoding UTF8
            Write-Host ("Saved {0} program ID(s) to {1}" -f $toSave.Count, $listPath)
        }

        "upgrade" {
            if (-not (Test-Path $listPath)) {
                $message = "Saved list file not found at $listPath. Run 'save' first."
                Write-Host $message
                Write-Log $message
                $script:ExitCode = 1
                break
            }
            Write-Host "Upgrading programs from saved list at $listPath..."
            $ids = @(Get-CleanLines -Path $listPath | Where-Object { $_ -notin $exceptions })
            Invoke-UpgradeSet -Ids $ids
        }

        "upgrade-all" {
            Write-Host "Resolving upgradeable programs..."
            $items = Get-UpgradeableItems -Source $Source
            $ids   = @($items | Where-Object { $_.Id -notin $exceptions } | ForEach-Object { $_.Id })
            Invoke-UpgradeSet -Ids $ids
        }
    }
}
catch {
    $message = "An error occurred: $($_.Exception.Message)"
    Write-Host $message -ForegroundColor Red
    Write-Log $message
    $script:ExitCode = 1
}

exit $script:ExitCode
