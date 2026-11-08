#Requires -Version 5.1

function ConvertTo-WindowsCommandLineArgument {
    [CmdletBinding()]
    param(
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Value
    )

    if ($null -eq $Value) {
        $Value = ''
    }

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('"')
    $backslashCount = 0

    foreach ($character in $Value.ToCharArray()) {
        if ($character -eq '\') {
            $backslashCount++
            continue
        }

        if ($character -eq '"') {
            if ($backslashCount -gt 0) {
                [void]$builder.Append(('\' * ($backslashCount * 2)))
            }
            [void]$builder.Append('\"')
            $backslashCount = 0
            continue
        }

        if ($backslashCount -gt 0) {
            [void]$builder.Append(('\' * $backslashCount))
            $backslashCount = 0
        }
        [void]$builder.Append($character)
    }

    if ($backslashCount -gt 0) {
        [void]$builder.Append(('\' * ($backslashCount * 2)))
    }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function Test-UacCancellation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Exception]$Exception
    )

    $candidate = $Exception
    while ($null -ne $candidate) {
        if ($candidate -is [System.ComponentModel.Win32Exception] -and $candidate.NativeErrorCode -eq 1223) {
            return $true
        }
        $candidate = $candidate.InnerException
    }

    return $false
}

function Invoke-OOShutUp10Elevation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$BatchPath,

        [AllowEmptyString()]
        [string]$CancellationMarkerPath,

        [AllowEmptyCollection()]
        [string[]]$InstallerArguments = @()
    )

    $batchPath = [System.IO.Path]::GetFullPath($BatchPath)
    if (-not (Test-Path -LiteralPath $batchPath -PathType Leaf)) {
        [Console]::Error.WriteLine('The installer batch file could not be found: {0}' -f $batchPath)
        return 1
    }

    $quotedArguments = @(
        foreach ($argument in $InstallerArguments) {
            ConvertTo-WindowsCommandLineArgument -Value $argument
        }
    )
    $argumentList = [string]::Join(' ', [string[]]$quotedArguments)
    $workingDirectory = [Environment]::CurrentDirectory

    $markerName = 'MYTECH_INSTALL_ELEVATION_ATTEMPTED'
    $previousMarker = [Environment]::GetEnvironmentVariable($markerName, 'Process')
    [Environment]::SetEnvironmentVariable($markerName, '1', 'Process')

    try {
        $process = Start-Process -FilePath $batchPath `
            -ArgumentList $argumentList `
            -Verb RunAs `
            -WorkingDirectory $workingDirectory `
            -Wait `
            -PassThru `
            -ErrorAction Stop

        if ($null -eq $process) {
            [Console]::Error.WriteLine('The elevated installer process did not return an exit status.')
            return 1
        }

        return [int]$process.ExitCode
    }
    catch {
        if (Test-UacCancellation -Exception $_.Exception) {
            if ([string]::IsNullOrWhiteSpace($CancellationMarkerPath)) {
                [Console]::Error.WriteLine('UAC elevation was cancelled.')
            }
            else {
                try {
                    [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($CancellationMarkerPath), 'UAC_CANCELLED')
                }
                catch {
                    [Console]::Error.WriteLine('UAC elevation was cancelled; the status marker could not be written: {0}' -f $_.Exception.Message)
                }
            }
            return 1223
        }

        [Console]::Error.WriteLine('Could not start the elevated installer: {0}' -f $_.Exception.Message)
        return 1
    }
    finally {
        [Environment]::SetEnvironmentVariable($markerName, $previousMarker, 'Process')
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    if ($args.Count -lt 1) {
        [Console]::Error.WriteLine('The installer batch path was not supplied.')
        exit 2
    }

    $batchPath = [string]$args[0]
    $cancellationMarkerPath = ''
    $installerArguments = @()
    if ($args.Count -gt 1) {
        $cancellationMarkerPath = [string]$args[1]
    }
    if ($args.Count -gt 2) {
        $installerArguments = [string[]]$args[2..($args.Count - 1)]
    }

    exit (Invoke-OOShutUp10Elevation -BatchPath $batchPath -CancellationMarkerPath $cancellationMarkerPath -InstallerArguments $installerArguments)
}
