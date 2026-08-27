$ErrorActionPreference = "Stop"

$executable = Resolve-Path "$PSScriptRoot\..\dist\DriveAutomate.exe"
$env:QT_QPA_PLATFORM = "offscreen"

function Invoke-SmokeTest {
    param(
        [Parameter(Mandatory)]
        [string] $Argument,

        [Parameter(Mandatory)]
        [string] $Name
    )

    $process = Start-Process `
        -FilePath $executable `
        -ArgumentList $Argument `
        -PassThru `
        -WindowStyle Hidden

    if (-not $process.WaitForExit(30000)) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        throw "$Name smoke test timed out after 30 seconds."
    }

    if ($process.ExitCode -ne 0) {
        throw "$Name smoke test failed with exit code $($process.ExitCode)."
    }

    Write-Host "$Name smoke test passed."
}

try {
    Invoke-SmokeTest -Argument "--help" -Name "CLI"
    Invoke-SmokeTest -Argument "--smoke-test-gui" -Name "PySide6 GUI"
}
finally {
    Remove-Item Env:QT_QPA_PLATFORM -ErrorAction SilentlyContinue
}
