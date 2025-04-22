param (
    [ValidateSet("list", "save", "upgrade")]
    [string] $command,
    [string] $listPath = "$PSScriptRoot\UpgradeablePrograms.txt",
    [string] $exceptionsPath = "$PSScriptRoot\PermanentUpgradeExceptions.txt"
)

if (-not $command) {
    Write-Host "Usage:"
    Write-Host "  .\WingetUpgrade.ps1 command <list|save|upgrade> -listPath <path> -exceptionsPath <path>"
    Write-Host ""
    Write-Host "Commands:"
    Write-Host "  list    - Lists all upgradeable programs."
    Write-Host "  save    - Saves upgradeable programs to a file, excluding exceptions."
    Write-Host "  upgrade - Upgrades programs from the saved list, excluding exceptions."
    Write-Host ""
    Write-Host "Optional Parameters:"
    Write-Host "  -listPath         Path to save or read the list of upgradeable programs. Defaults to $PSScriptRoot\\UpgradeablePrograms.txt."
    Write-Host "  -exceptionsPath   Path to a file containing program IDs to exclude from upgrades. Defaults to $PSScriptRoot\\PermanentUpgradeExceptions.txt."
    exit
}

function Get-WingetUpgrades {
    # Run the winget upgrade command and capture raw output with UTF-8 encoding
    Start-Process -FilePath "winget" -ArgumentList "upgrade" -NoNewWindow -Wait -PassThru -RedirectStandardOutput "winget_output.txt"
    $output = Get-Content -Path "winget_output.txt" -Encoding UTF8
    Remove-Item -Path "winget_output.txt" -Force
    return $output
}

function Convert-WingetOutput {
    param (
        [string[]] $OutputLines
    )

    $results = @()
    $headerIndex = -1
    $columns = @{}

    # Locate the header line and extract column positions
    for ($i = 0; $i -lt $OutputLines.Length; $i++) {
        $line = $OutputLines[$i]

        if ($line -match "Name\s+Id\s+Version\s+Available\s+Source") {
            $headerIndex = $i

            # Extract column positions based on header
            $columns = @{
                NameColumn      = $line.IndexOf("Name")
                IdColumn        = $line.IndexOf("Id")
                VersionColumn   = $line.IndexOf("Version")
                AvailableColumn = $line.IndexOf("Available")
                SourceColumn    = $line.IndexOf("Source")
            }
            break
        }
    }

    # Process rows after the header
    for ($i = $headerIndex + 2; $i -lt $OutputLines.Length; $i++) {
        $line = $OutputLines[$i]

        # Stop at summary line
        if ($line -match '^\d+ upgrades available') {
            break
        }

        # Skip empty lines
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        # Extract fields using column positions
        $results += [PSCustomObject]@{
            Name      = $line.Substring($columns.NameColumn, $columns.IdColumn - $columns.NameColumn).Trim()
            Id        = $line.Substring($columns.IdColumn, $columns.VersionColumn - $columns.IdColumn).Trim()
            Version   = $line.Substring($columns.VersionColumn, $columns.AvailableColumn - $columns.VersionColumn).Trim()
            Available = $line.Substring($columns.AvailableColumn, $columns.SourceColumn - $columns.AvailableColumn).Trim()
            Source    = $line.Substring($columns.SourceColumn).Trim()
        }
    }

    return $results
}

function Get-UpgradeablePrograms {
    param (
        [string[]] $OutputLines
    )
    
    Write-Host "Listing upgradeable programs..."
    $programs = Convert-WingetOutput -OutputLines $OutputLines
    
    if ($programs.Count -eq 0) {
        Write-Host "No upgrades available or unable to capture IDs."
    }
    else {
        $programs | ForEach-Object { 
            Write-Host "Program: $($_.Name), ID: $($_.Id)" 
        }
    }
    
    return $programs
}

function Save-UpgradeablePrograms {
    param (
        [string[]] $OutputLines,
        [string] $FilePath,
        [string[]] $Exceptions
    )
    
    Write-Host "Saving upgradeable programs to $FilePath..."
    $programs = Convert-WingetOutput -OutputLines $OutputLines

    if ($programs.Count -eq 0) {
        Write-Host "No upgrades available or unable to capture IDs."
    }
    else {
        $programs | Where-Object { $_.Id -notin $Exceptions } | ForEach-Object {
            Write-Host "Saving Program: $($_.Name), ID: $($_.Id)"
            $_.Id
        } | Out-File -FilePath $FilePath
        Write-Host "Upgradeable programs saved to $FilePath"
    }
}

function Invoke-FromSavedList {
    param (
        [string] $FilePath,
        [string[]] $Exceptions
    )
    
    if (Test-Path $FilePath) {
        Write-Host "Upgrading programs from saved list at $FilePath..."
        Get-Content -Path $FilePath | Where-Object { $_ -notin $Exceptions } | ForEach-Object {
            Write-Host "Upgrading Program ID: $_"
            try {
                winget upgrade $_
            }
            catch {
                $errorMessage = "Error upgrading Program ID: $_ - $($_.Exception.Message)"
                Write-Host $errorMessage
                Add-Content -Path "WingetUpgradeErrors.log" -Value "$(Get-Date): $errorMessage"
            }
        }
    }
    else {
        Write-Host "Saved list file not found."
        Add-Content -Path "WingetUpgradeErrors.log" -Value "$(Get-Date): Saved list file not found at $FilePath."
    }
}

# Main script execution
try {
    $wingetOutput = Get-WingetUpgrades
    $outputLines = $wingetOutput -split "`r?`n"
    $exceptions = @()
    if (Test-Path $exceptionsPath) {
        $exceptions = Get-Content -Path $exceptionsPath -ErrorAction SilentlyContinue
    }

    switch ($command) {
        "list" { Get-UpgradeablePrograms -OutputLines $outputLines }
        "save" { Save-UpgradeablePrograms -OutputLines $outputLines -FilePath $listPath -Exceptions $exceptions }
        "upgrade" { Invoke-FromSavedList -FilePath $listPath -Exceptions $exceptions }
        default { Write-Host "Invalid command provided. Use 'list', 'save', or 'upgrade'." }
    }
}
catch {
    Write-Host "An error occurred: $($_.Exception.Message)"
}
finally {
    if (Test-Path "winget_output.txt") {
        Remove-Item -Path "winget_output.txt" -Force
    }
}