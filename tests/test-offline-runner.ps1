$ErrorActionPreference = 'Stop'

$runnerPath = Join-Path $PSScriptRoot 'run-offline-tests.ps1'
$engineName = if ($PSVersionTable.PSEdition -eq 'Core') {
    if ($env:OS -eq 'Windows_NT') { 'pwsh.exe' } else { 'pwsh' }
}
else {
    'powershell.exe'
}
$enginePath = Join-Path $PSHOME $engineName

if (-not (Test-Path -LiteralPath $runnerPath -PathType Leaf)) {
    throw 'The offline runner script is missing.'
}
if (-not (Test-Path -LiteralPath $enginePath -PathType Leaf)) {
    throw 'The current PowerShell executable is unavailable.'
}

$fixtureRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('offline-runner-failure-' + [guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($fixtureRoot)

try {
    $passingCommandPath = Join-Path $fixtureRoot 'fake-passing-command.ps1'
    $fakeCommandPath = Join-Path $fixtureRoot 'fake-failing-command.ps1'
    [System.IO.File]::WriteAllText(
        $passingCommandPath,
        "Write-Output 'fake passing marker'`r`nexit 0`r`n",
        [System.Text.Encoding]::UTF8
    )
    [System.IO.File]::WriteAllText(
        $fakeCommandPath,
        "Write-Output 'fake failure marker'`r`nexit 23`r`n",
        [System.Text.Encoding]::UTF8
    )

    $commandFixturePath = Join-Path $fixtureRoot 'commands.json'
    $commandFixturesJson = @{
        Commands = @(
            [pscustomobject]@{
                Label = 'fake passing command'
                Path = $passingCommandPath
            },
            [pscustomobject]@{
                Label = 'fake failing command'
                Path = $fakeCommandPath
            }
        )
    } | ConvertTo-Json -Depth 4 -Compress
    [System.IO.File]::WriteAllText($commandFixturePath, $commandFixturesJson, [System.Text.Encoding]::UTF8)

    $runnerArguments = @(
        '-NoLogo',
        '-NoProfile',
        '-NonInteractive',
        '-File',
        $runnerPath,
        '-CommandFixturePath',
        $commandFixturePath
    )
    $capturedOutput = & $enginePath @runnerArguments 2>&1
    $exitCode = $LASTEXITCODE
    $outputText = [string]::Join("`n", [string[]]@($capturedOutput | ForEach-Object { [string]$_ }))

    if ($exitCode -ne 23) {
        throw ("The aggregate runner returned {0}; expected the fake command's exit code 23." -f $exitCode)
    }
    if (-not $outputText.Contains('PASS: fake passing command')) {
        throw 'The aggregate runner did not report the passing fixture.'
    }
    if (-not $outputText.Contains('FAIL: fake failing command (exit 23)')) {
        throw 'The aggregate runner did not identify the fake failing command and its exit status.'
    }
    if (-not $outputText.Contains('fake passing marker')) {
        throw 'The aggregate runner did not execute the passing fixture.'
    }
    if (-not $outputText.Contains('fake failure marker')) {
        throw 'The aggregate runner did not execute the fake failing command.'
    }

    Write-Output 'PASS: runner failure propagation preserves exit 23 and identifies the command.'
}
finally {
    $temporaryRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    ) + [System.IO.Path]::DirectorySeparatorChar
    $resolvedFixtureRoot = [System.IO.Path]::GetFullPath($fixtureRoot)
    $comparison = if ($env:OS -eq 'Windows_NT') {
        [System.StringComparison]::OrdinalIgnoreCase
    }
    else {
        [System.StringComparison]::Ordinal
    }

    if (-not $resolvedFixtureRoot.StartsWith($temporaryRoot, $comparison)) {
        throw 'The runner fixture cleanup path is outside the system temporary directory.'
    }

    if (Test-Path -LiteralPath $resolvedFixtureRoot) {
        Remove-Item -LiteralPath $resolvedFixtureRoot -Recurse -Force
    }
}
