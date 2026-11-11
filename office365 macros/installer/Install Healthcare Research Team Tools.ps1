# Healthcare Research Team Tools installer/updater
# Installs the global Word template for the current Windows user.
# No administrator rights are required.

$ErrorActionPreference = "Stop"

$ProductName = "Healthcare Research Team Tools"
$PackageVersion = "1.0.0"
$TemplateFileName = "Healthcare Research Team Tools.dotm"

$ScriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
$SourceTemplate = Join-Path $ScriptFolder $TemplateFileName

$WordFolder = Join-Path $env:APPDATA "Microsoft\Word"
$StartupFolder = Join-Path $WordFolder "STARTUP"
$DestinationTemplate = Join-Path $StartupFolder $TemplateFileName

$TemporaryTemplate = Join-Path `
    $StartupFolder `
    ($TemplateFileName + ".installing")

$BackupTemplate = Join-Path `
    $StartupFolder `
    ($TemplateFileName + ".previous")

function Show-Message {
    param(
        [string]$Message,
        [string]$Title = $ProductName,
        [int]$Icon = 64
    )

    $shell = New-Object -ComObject WScript.Shell
    $null = $shell.Popup($Message, 0, $Title, $Icon)
}

function Stop-WithError {
    param([string]$Message)

    Show-Message `
        -Message $Message `
        -Title "$ProductName - Installation Error" `
        -Icon 16

    exit 1
}

try {
    if (-not (Test-Path -LiteralPath $SourceTemplate)) {
        Stop-WithError @"
The installer could not find:

$TemplateFileName

Keep the installer and template together in the same extracted folder, then run the installer again.
"@
    }

    $wordProcesses = Get-Process `
        -Name "WINWORD" `
        -ErrorAction SilentlyContinue

    if ($wordProcesses) {
        Stop-WithError @"
Microsoft Word is currently open.

Close every Word window, wait a few seconds, and run the installer again.

The installer will not close Word automatically because doing so could discard unsaved work.
"@
    }

    if (-not (Test-Path -LiteralPath $StartupFolder)) {
        $null = New-Item `
            -ItemType Directory `
            -Path $StartupFolder `
            -Force
    }

    # Remove the Mark of the Web from the downloaded source file.
    Unblock-File `
        -LiteralPath $SourceTemplate `
        -ErrorAction SilentlyContinue

    # Remove residue from an interrupted prior installation.
    if (Test-Path -LiteralPath $TemporaryTemplate) {
        Remove-Item `
            -LiteralPath $TemporaryTemplate `
            -Force
    }

    # Copy first to a temporary file in the destination folder.
    Copy-Item `
        -LiteralPath $SourceTemplate `
        -Destination $TemporaryTemplate `
        -Force

    Unblock-File `
        -LiteralPath $TemporaryTemplate `
        -ErrorAction SilentlyContinue

    # Verify the copied file before replacing the installed version.
    $sourceHash = Get-FileHash `
        -LiteralPath $SourceTemplate `
        -Algorithm SHA256

    $temporaryHash = Get-FileHash `
        -LiteralPath $TemporaryTemplate `
        -Algorithm SHA256

    if ($sourceHash.Hash -ne $temporaryHash.Hash) {
        Remove-Item `
            -LiteralPath $TemporaryTemplate `
            -Force `
            -ErrorAction SilentlyContinue

        Stop-WithError @"
The copied template did not pass verification.

No installed version was replaced. Download a fresh copy of the package and try again.
"@
    }

    $wasUpdate = Test-Path -LiteralPath $DestinationTemplate

    # Retain one temporary backup during replacement.
    if (Test-Path -LiteralPath $BackupTemplate) {
        Remove-Item `
            -LiteralPath $BackupTemplate `
            -Force
    }

    if ($wasUpdate) {
        Move-Item `
            -LiteralPath $DestinationTemplate `
            -Destination $BackupTemplate `
            -Force
    }

    try {
        Move-Item `
            -LiteralPath $TemporaryTemplate `
            -Destination $DestinationTemplate `
            -Force
    }
    catch {
        # Restore the prior installed copy if final replacement fails.
        if (Test-Path -LiteralPath $BackupTemplate) {
            Move-Item `
                -LiteralPath $BackupTemplate `
                -Destination $DestinationTemplate `
                -Force
        }

        throw
    }

    Unblock-File `
        -LiteralPath $DestinationTemplate `
        -ErrorAction SilentlyContinue

    # Verify the final installed file.
    $installedHash = Get-FileHash `
        -LiteralPath $DestinationTemplate `
        -Algorithm SHA256

    if ($sourceHash.Hash -ne $installedHash.Hash) {
        if (Test-Path -LiteralPath $DestinationTemplate) {
            Remove-Item `
                -LiteralPath $DestinationTemplate `
                -Force
        }

        if (Test-Path -LiteralPath $BackupTemplate) {
            Move-Item `
                -LiteralPath $BackupTemplate `
                -Destination $DestinationTemplate `
                -Force
        }

        Stop-WithError @"
The installed template did not pass verification.

The previous version was restored when available.
"@
    }

    if (Test-Path -LiteralPath $BackupTemplate) {
        Remove-Item `
            -LiteralPath $BackupTemplate `
            -Force
    }

    if ($wasUpdate) {
        $action = "updated"
    }
    else {
        $action = "installed"
    }

    Show-Message @"
$ProductName was $action successfully.

Version: $PackageVersion

Open Microsoft Word to use the updated tool.
"@

    exit 0
}
catch {
    if (Test-Path -LiteralPath $TemporaryTemplate) {
        Remove-Item `
            -LiteralPath $TemporaryTemplate `
            -Force `
            -ErrorAction SilentlyContinue
    }

    Stop-WithError @"
$ProductName could not be installed.

$($_.Exception.Message)

No administrator rights should be required. If the problem continues, send a screenshot of this message to Michael.
"@
}