$ErrorActionPreference = "Stop"

& (Join-Path $PSScriptRoot "stop.ps1")

$desktop = [Environment]::GetFolderPath("Desktop")
$startup = [Environment]::GetFolderPath("Startup")
$programs = [Environment]::GetFolderPath("Programs")
$startMenuFolder = Join-Path $programs "Keyme"
$legacyStartMenuFolder = Join-Path $programs "Keeby Windows"

$paths = @(
    (Join-Path $desktop "Start Keyme.lnk"),
    (Join-Path $desktop "Stop Keyme.lnk"),
    (Join-Path $desktop "Keyme.lnk"),
    (Join-Path $desktop "Start Keeby Windows.lnk"),
    (Join-Path $desktop "Stop Keeby Windows.lnk"),
    (Join-Path $startup "Keyme.lnk"),
    (Join-Path $startup "Keeby Windows.lnk"),
    (Join-Path $startMenuFolder "Start Keyme.lnk"),
    (Join-Path $startMenuFolder "Stop Keyme.lnk"),
    (Join-Path $startMenuFolder "Keyme.lnk")
)

foreach ($path in $paths) {
    Remove-Item $path -ErrorAction SilentlyContinue
}

Remove-Item $startMenuFolder -ErrorAction SilentlyContinue
Remove-Item $legacyStartMenuFolder -Recurse -ErrorAction SilentlyContinue
Remove-Item (Join-Path $env:LOCALAPPDATA "Keyme") -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path $env:APPDATA "Keyme") -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Removed Keyme shortcuts, settings, and installed app files."
