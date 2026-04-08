param(
    [string]$Version = ""
)

$ErrorActionPreference = "Stop"

$sourceRoot = "D:\Work\Tools\PremiereProjectCollector"
$archiveRoot = "D:\Work\Tools\PremiereProjectCollector_Versions"

if (-not $Version -or -not $Version.Trim()) {
    $versionFile = Join-Path $sourceRoot "version.json"
    if (-not (Test-Path $versionFile)) {
        throw "Could not find version.json at $versionFile"
    }

    $versionData = Get-Content $versionFile -Raw | ConvertFrom-Json
    $Version = [string]$versionData.version
}

if (-not $Version -or -not $Version.Trim()) {
    throw "Version is required."
}

$destination = Join-Path $archiveRoot $Version

New-Item -ItemType Directory -Force $destination | Out-Null

$excludeDirs = @(".git")
$excludeFiles = @("archive_release.ps1")

$robocopyArgs = @(
    $sourceRoot,
    $destination,
    "/MIR"
)

foreach ($dir in $excludeDirs) {
    $robocopyArgs += "/XD"
    $robocopyArgs += $dir
}

foreach ($file in $excludeFiles) {
    $robocopyArgs += "/XF"
    $robocopyArgs += $file
}

& robocopy @robocopyArgs | Out-Null

if ($LASTEXITCODE -ge 8) {
    throw "Robocopy failed with exit code $LASTEXITCODE"
}

$releaseInfo = @{
    version = $Version
    source = $sourceRoot
    archivedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
} | ConvertTo-Json

Set-Content -Path (Join-Path $destination "release-info.json") -Value $releaseInfo -Encoding UTF8

Write-Host "Archived Project Collector version $Version to $destination"
