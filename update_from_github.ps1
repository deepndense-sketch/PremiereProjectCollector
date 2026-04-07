param(
    [string]$ZipPath = "",
    [string]$Destination = "",
    [string]$ResultPath = "",
    [string]$LogPath = ""
)

$ErrorActionPreference = "Stop"

if (-not $LogPath -or -not $LogPath.Trim()) {
    $LogPath = Join-Path $env:TEMP "PremiereProjectCollector_update_log.txt"
}

if (-not $ResultPath -or -not $ResultPath.Trim()) {
    $ResultPath = Join-Path $env:TEMP "PremiereProjectCollector_update_result.json"
}

if (-not $Destination -or -not $Destination.Trim()) {
    $Destination = Join-Path $env:APPDATA "Adobe\CEP\extensions\PremiereProjectCollector"
}

function Write-Step($message) {
    Write-Host "[Project Collector Updater] $message"
    Add-Content -Path $LogPath -Value ("[STEP] " + $message)
}

function Write-Result($ok, $message) {
    $payload = @{
        ok = $ok
        message = $message
        logPath = $LogPath
        zipPath = $ZipPath
        destination = $Destination
    } | ConvertTo-Json -Compress

    Set-Content -Path $ResultPath -Value $payload -Encoding UTF8
}

$tempRoot = Join-Path $env:TEMP ("PremiereProjectCollectorUpdate_" + [guid]::NewGuid().ToString("N"))
$extractPath = Join-Path $tempRoot "extract"

try {
    Set-Content -Path $LogPath -Value ("Project Collector updater started at " + (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) -Encoding UTF8
    Write-Step "Preparing temporary workspace."
    New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null
    New-Item -ItemType Directory -Force -Path $extractPath | Out-Null

    if (-not $ZipPath -or -not (Test-Path $ZipPath)) {
        throw "Prepared update package was not found: $ZipPath"
    }

    Write-Step "Using prepared package: $ZipPath"

    Write-Step "Extracting package."
    Expand-Archive -Path $ZipPath -DestinationPath $extractPath -Force

    $sourceRoot = Join-Path $extractPath "PremiereProjectCollector-main"
    if (-not (Test-Path $sourceRoot)) {
        throw "Could not find extracted PremiereProjectCollector-main folder."
    }

    Write-Step "Ensuring CEP extensions destination exists."
    New-Item -ItemType Directory -Force -Path $Destination | Out-Null

    Write-Step "Copying updated extension files."
    $robocopyOutput = robocopy $sourceRoot $Destination /MIR /XD .git
    $robocopyExit = $LASTEXITCODE
    if ($robocopyExit -ge 8) {
        throw "Robocopy failed with exit code $robocopyExit."
    }

    Write-Step "Update completed successfully."
    Write-Host "Restart Premiere Pro if the panel is already open."
    Write-Result $true "Update completed successfully."
}
catch {
    $message = $_.Exception.Message
    Add-Content -Path $LogPath -Value ("[ERROR] " + $message)
    Write-Host "[Project Collector Updater] ERROR: $message"
    Write-Host "[Project Collector Updater] Log file: $LogPath"
    Write-Result $false $message
    Read-Host "Update failed. Press Enter to close this window"
}
finally {
    if (Test-Path $tempRoot) {
        Remove-Item -Recurse -Force $tempRoot -ErrorAction SilentlyContinue
    }
}
