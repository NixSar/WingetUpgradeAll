<#
.SYNOPSIS
    Lists, saves, and bulk-upgrades programs via the Windows Package Manager (winget),
    with support for a permanent exceptions list and scope-aware elevation.

.DESCRIPTION
    WingetUpgradeAll.ps1 wraps winget to make unattended bulk upgrades easy.

    Commands:
      list         Show all upgradeable programs (with install scope).
      save         Save upgradeable program IDs to a file, excluding exceptions.
      upgrade      Upgrade every program listed in a previously saved file.
      upgrade-all  List live and upgrade everything in one step, excluding exceptions.

    When the Microsoft.WinGet.Client PowerShell module is installed it is used to
    enumerate upgradeable packages (robust, locale-independent). Otherwise the script
    falls back to parsing `winget upgrade` text output. Upgrades themselves always run
    through the winget CLI so failures are detected via the process exit code.

    Scope-aware elevation (2.1):
      Per-user packages (installed under the user profile, e.g. Electron/Squirrel
      apps) must NOT be upgraded from an elevated process: their installers then
      write into the user hive with an elevated token and leave behind
      Administrators-owned registry keys with wrong permissions, which can break
      other installers (e.g. MSIX file-type registration).

      The script determines each package's install scope via `winget list --scope`
      and runs two passes so that every package is upgraded under the right token,
      whichever kind of terminal it was started from:

        Started elevated:     machine-scope packages upgrade in-process; user-scope
                              and unknown-scope packages run in ONE de-elevated
                              child (launched via the desktop shell, no prompt).
        Started non-elevated: user-scope and unknown-scope packages upgrade
                              in-process; machine-scope packages run in ONE
                              elevated child (one UAC prompt per run).

.PARAMETER command
    Operation to perform: list, save, upgrade, or upgrade-all.
    (_batch is internal: used for the child runs described above.)

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
    Show what would be upgraded (and in which scope/pass) without performing any upgrades.

.PARAMETER resultPath
    Internal. Result file written by a child run.

.PARAMETER pidPath
    Internal. File a child run writes its process ID to on start-up.

.EXAMPLE
    .\WingetUpgradeAll.ps1 list

.EXAMPLE
    .\WingetUpgradeAll.ps1 save -exceptionsPath .\Exceptions.txt

.EXAMPLE
    .\WingetUpgradeAll.ps1 upgrade

.EXAMPLE
    .\WingetUpgradeAll.ps1 upgrade-all -Source winget -WhatIf

.NOTES
    Version: 2.1.0
#>
#Requires -Version 5.1

param (
    [Parameter(Position = 0)]
    [ValidateSet("list", "save", "upgrade", "upgrade-all", "_batch")]
    [string] $command,

    [string] $listPath       = "$PSScriptRoot\UpgradeablePrograms.txt",
    [string] $exceptionsPath  = "$PSScriptRoot\PermanentUpgradeExceptions.txt",
    [string] $logPath         = "$PSScriptRoot\WingetUpgradeErrors.log",

    [ValidateSet("winget", "msstore")]
    [string] $Source,

    [switch] $Interactive,
    [switch] $WhatIf,

    # Internal (child runs)
    [string] $resultPath,
    [string] $pidPath
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

function Test-Elevated {
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
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

function ConvertFrom-WingetTableText {
    # Parses the table printed by `winget upgrade` and `winget list`
    # (both use the columns Name / Id / Version / Available / Source).
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

function Get-WingetText {
    # Run winget with the given arguments and return its stdout lines.
    # Output goes to a private temp file (avoids CWD clutter and name
    # collisions between concurrent runs).
    param ([string[]] $Arguments)

    $tempFile = Join-Path $env:TEMP ("winget_{0}_{1}.txt" -f $PID, [guid]::NewGuid().ToString('N'))
    try {
        $null = Start-Process -FilePath 'winget' `
            -ArgumentList $Arguments `
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
    $items = ConvertFrom-WingetTableText -OutputLines (Get-WingetText -Arguments @('upgrade'))
    if ($Source) {
        $items = $items | Where-Object { $_.Source -eq $Source }
    }
    return $items
}

function Get-PackageScopeMap {
    # Returns a hashtable Id -> 'user' | 'machine', built from
    # `winget list --scope user` and `winget list --scope machine`.
    # IDs found in both scopes are treated as 'machine' (the machine copy needs
    # elevation). IDs found in neither are absent from the map ('unknown').
    $map = @{}

    foreach ($scope in 'user', 'machine') {
        $lines = Get-WingetText -Arguments @('list', '--scope', $scope, '--accept-source-agreements')
        $items = ConvertFrom-WingetTableText -OutputLines $lines
        if (@($items).Count -eq 0) {
            Write-Host "  Warning: 'winget list --scope $scope' returned no parsable packages." -ForegroundColor Yellow
        }
        foreach ($item in $items) {
            if ($scope -eq 'machine' -or -not $map.ContainsKey($item.Id)) {
                $map[$item.Id] = $scope
            }
        }
    }

    return $map
}

function Get-PackageScope {
    param ([hashtable] $ScopeMap, [string] $Id)
    if ($ScopeMap.ContainsKey($Id)) { return $ScopeMap[$Id] }
    return 'unknown'
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

function Close-ConsoleLine {
    # winget's progress rendering can leave the cursor mid-line; close that
    # line first so the separator below is actually blank.
    try {
        if ($Host.UI.RawUI.CursorPosition.X -gt 0) { Write-Host "" }
    } catch {}
    Write-Host ""
}

function New-BatchContext {
    # Temp files and child argument list shared by both child-run launchers.
    param ([string[]] $Ids)

    $token = [guid]::NewGuid().ToString('N')
    $ctx = @{
        List   = Join-Path $env:TEMP "winget_batch_$token.txt"
        Result = Join-Path $env:TEMP "winget_result_$token.txt"
        Pid    = Join-Path $env:TEMP "winget_pid_$token.txt"
    }
    $Ids | Set-Content -Path $ctx.List -Encoding UTF8

    $childArgs = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $PSCommandPath),
        '_batch',
        '-listPath',   ('"{0}"' -f $ctx.List),
        '-resultPath', ('"{0}"' -f $ctx.Result),
        '-pidPath',    ('"{0}"' -f $ctx.Pid),
        '-logPath',    ('"{0}"' -f $logPath)
    )
    if ($Interactive) { $childArgs += '-Interactive' }
    $ctx.Args    = $childArgs
    $ctx.HostExe = (Get-Process -Id $PID).Path
    return $ctx
}

function Remove-BatchContext {
    param ([hashtable] $Ctx)
    foreach ($f in $Ctx.List, $Ctx.Result, $Ctx.Pid) {
        if ($f -and (Test-Path $f)) { Remove-Item -Path $f -Force -ErrorAction SilentlyContinue }
    }
}

function Read-BatchResult {
    # Merge the child's "Id<TAB>Status" lines into a hashtable; anything the
    # child did not report is a failure (child crashed, was killed, etc.).
    param ([hashtable] $Ctx, [string[]] $Ids, [string] $Kind, [string] $Detail)

    $statuses = @{}
    if (Test-Path $Ctx.Result) {
        foreach ($line in Get-Content -Path $Ctx.Result -Encoding UTF8) {
            $parts = $line -split "`t", 2
            if ($parts.Count -ne 2) { continue }
            if ($parts[0] -eq '#elevated') {
                $childElevated = [bool]::Parse($parts[1])
                if ($childElevated -ne ($Kind -eq 'elevated')) {
                    Write-Host ("  Warning: {0} child ran with elevated={1}." -f $Kind, $childElevated) -ForegroundColor Yellow
                }
                continue
            }
            $statuses[$parts[0]] = $parts[1]
        }
    }
    foreach ($id in $Ids) {
        if (-not $statuses.ContainsKey($id)) {
            $message = "$Kind run returned no result for '$id'$Detail."
            Write-Host "  $message" -ForegroundColor Red
            Write-Log $message
            $statuses[$id] = 'Failed'
        }
    }
    return $statuses
}

function Invoke-ElevatedBatch {
    # Upgrade the given machine-scope IDs in ONE elevated child run of this
    # script (single UAC prompt). Returns a hashtable Id -> status.
    param ([string[]] $Ids)

    if (-not $Ids -or @($Ids).Count -eq 0) { return @{} }
    $ctx = New-BatchContext -Ids $Ids
    try {
        Write-Host ("Elevating once for {0} machine-scope package(s)..." -f @($Ids).Count)
        try {
            $proc = Start-Process -FilePath $ctx.HostExe -ArgumentList $ctx.Args `
                -Verb RunAs -Wait -PassThru
        }
        catch {
            # UAC declined or elevation otherwise failed.
            $message = "Elevation declined or failed; machine-scope packages were not upgraded ($($_.Exception.Message))."
            Write-Host "  $message" -ForegroundColor Red
            Write-Log $message
            $statuses = @{}
            foreach ($id in $Ids) { $statuses[$id] = 'Failed' }
            return $statuses
        }
        return Read-BatchResult -Ctx $ctx -Ids $Ids -Kind 'elevated' -Detail " (child exit code $($proc.ExitCode))"
    }
    finally {
        Remove-BatchContext -Ctx $ctx
    }
}

function Invoke-DeElevatedBatch {
    # Upgrade the given user-scope IDs in ONE non-elevated child run of this
    # script, launched from an elevated parent via the desktop shell (Explorer),
    # which runs at the user's normal medium integrity level. No prompt.
    # (The linked-token/CreateProcessWithTokenW route is not usable without
    # SeTcbPrivilege; ShellExecute through the desktop is the reliable one.)
    # ShellExecute returns no handle, so the child writes its PID on start-up
    # and we wait on that process. Returns a hashtable Id -> status.
    param ([string[]] $Ids)

    if (-not $Ids -or @($Ids).Count -eq 0) { return @{} }
    $ctx = New-BatchContext -Ids $Ids
    try {
        Write-Host ("De-elevating once for {0} user-scope/unknown package(s)..." -f @($Ids).Count)
        try {
            $shellWindows = [Activator]::CreateInstance([type]::GetTypeFromCLSID('9BA05972-F6A8-11CF-A442-00A0C90A8F39'))
            $desktopShell = $shellWindows.Item().Document.Application
            $desktopShell.ShellExecute($ctx.HostExe, ($ctx.Args -join ' '), '', 'open', 1)
        }
        catch {
            $message = "Could not start the non-elevated child (is Explorer running?): $($_.Exception.Message)"
            Write-Host "  $message" -ForegroundColor Red
            Write-Log $message
            $statuses = @{}
            foreach ($id in $Ids) { $statuses[$id] = 'Failed' }
            return $statuses
        }

        # Wait for the child to announce itself, then for it to finish.
        $deadline = (Get-Date).AddSeconds(60)
        $childPid = $null
        while (-not $childPid -and (Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 250
            if (Test-Path $ctx.Pid) {
                $raw = (Get-Content -Path $ctx.Pid -ErrorAction SilentlyContinue | Select-Object -First 1)
                if ($raw -match '^\d+$') { $childPid = [int]$raw }
            }
        }
        if (-not $childPid) {
            return Read-BatchResult -Ctx $ctx -Ids $Ids -Kind 'de-elevated' -Detail ' (child never started)'
        }
        try { Wait-Process -Id $childPid -ErrorAction Stop } catch {}   # already exited is fine
        return Read-BatchResult -Ctx $ctx -Ids $Ids -Kind 'de-elevated' -Detail ''
    }
    finally {
        Remove-BatchContext -Ctx $ctx
    }
}

function Invoke-UpgradeSet {
    param ([string[]] $Ids)

    if (-not $Ids -or @($Ids).Count -eq 0) {
        Write-Host "Nothing to upgrade."
        return
    }

    $elevated = Test-Elevated

    Write-Host "Determining install scope of packages..."
    $scopeMap = Get-PackageScopeMap

    $userIds    = @()   # user-scope + unknown: run in-process, never elevated
    $machineIds = @()   # machine-scope: elevated pass
    foreach ($id in $Ids) {
        if ((Get-PackageScope -ScopeMap $scopeMap -Id $id) -eq 'machine') { $machineIds += $id }
        else                                                              { $userIds    += $id }
    }

    $tally = @{ Succeeded = 0; Failed = 0; Skipped = 0 }

    if ($WhatIf) {
        Write-Host ""
        Write-Host ("[WhatIf] Non-elevated pass ({0}):" -f $userIds.Count)
        foreach ($id in $userIds) {
            Write-Host ("  {0}  [{1}]" -f $id, (Get-PackageScope -ScopeMap $scopeMap -Id $id))
        }
        Write-Host ("[WhatIf] Elevated pass ({0}):" -f $machineIds.Count)
        foreach ($id in $machineIds) { Write-Host "  $id  [machine]" }
        Write-Host ""
        Write-Host ("Summary: 0 succeeded, 0 failed, {0} skipped." -f $Ids.Count)
        return
    }

    # ---- Pass 1: user-scope / unknown (must NOT run elevated) ----
    if ($userIds.Count -gt 0) {
        Write-Host ""
        if ($elevated) {
            $statuses = Invoke-DeElevatedBatch -Ids $userIds
            foreach ($id in $userIds) {
                $status = $statuses[$id]
                Write-Host ("  {0}: {1}" -f $status, $id)
                $tally[$status]++
            }
        }
        else {
            Write-Host ("Non-elevated pass: {0} package(s)" -f $userIds.Count)
            foreach ($id in $userIds) {
                $status = Invoke-Upgrade -Id $id -Interactive:$Interactive
                $tally[$status]++
                Close-ConsoleLine
            }
        }
    }

    # ---- Pass 2: machine-scope (must run elevated) ----
    if ($machineIds.Count -gt 0) {
        Write-Host ""
        if ($elevated) {
            Write-Host ("Elevated pass: {0} package(s)" -f $machineIds.Count)
            foreach ($id in $machineIds) {
                $status = Invoke-Upgrade -Id $id -Interactive:$Interactive
                $tally[$status]++
                Close-ConsoleLine
            }
        }
        else {
            $statuses = Invoke-ElevatedBatch -Ids $machineIds
            foreach ($id in $machineIds) {
                $status = $statuses[$id]
                Write-Host ("  {0}: {1}" -f $status, $id)
                $tally[$status]++
            }
        }
    }
    Write-Host ""

    Write-Host ("Summary: {0} succeeded, {1} failed, {2} skipped." -f `
        $tally.Succeeded, $tally.Failed, $tally.Skipped)

    if ($tally.Failed -gt 0) {
        $script:ExitCode = 1
    }
}

function Invoke-BatchChild {
    # Runs inside a child (elevated or de-elevated): announce PID, upgrade every
    # ID in $listPath, write "Id<TAB>Status" lines to $resultPath after each
    # package (so a crash mid-way still reports the ones done). Failures are
    # logged here (the child has the exit codes); the parent only tallies.
    if (-not $resultPath) {
        Write-Host "_batch requires -resultPath." -ForegroundColor Red
        exit 2
    }
    if ($pidPath) { Set-Content -Path $pidPath -Value $PID -Encoding ASCII }

    $elevated = Test-Elevated
    Write-Host ("Child run (elevated={0}): upgrading from {1}" -f $elevated, $listPath)
    "#elevated`t$elevated" | Set-Content -Path $resultPath -Encoding UTF8

    $ids = @(Get-CleanLines -Path $listPath)
    foreach ($id in $ids) {
        $status = Invoke-Upgrade -Id $id -Interactive:$Interactive
        "{0}`t{1}" -f $id, $status | Add-Content -Path $resultPath -Encoding UTF8
        Close-ConsoleLine
    }
}

function Show-Usage {
    Write-Host "Usage:"
    Write-Host "  .\WingetUpgradeAll.ps1 <list|save|upgrade|upgrade-all> [options]"
    Write-Host ""
    Write-Host "Scope-aware: machine-scope packages are upgraded elevated, user-scope"
    Write-Host "packages non-elevated, whichever kind of terminal you start from."
    Write-Host "(Elevated terminal: user-scope runs in a de-elevated child, no prompt."
    Write-Host " Normal terminal: machine-scope runs in one elevated child, one UAC prompt.)"
    Write-Host ""
    Write-Host "Commands:"
    Write-Host "  list         Lists all upgradeable programs with install scope."
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
    Write-Host "  -WhatIf                 Show what would be upgraded (and in which pass) without doing it."
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

if ($command -eq '_batch') {
    try {
        Invoke-BatchChild
    }
    catch {
        Write-Log "Batch child error: $($_.Exception.Message)"
        exit 1
    }
    exit 0
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
            $scopeMap = Get-PackageScopeMap
            $items |
                Select-Object Name, Id, Version, Available, Source,
                    @{ Name = 'Scope'; Expression = { Get-PackageScope -ScopeMap $scopeMap -Id $_.Id } } |
                Format-Table Name, Id, Version, Available, Source, Scope -AutoSize
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
