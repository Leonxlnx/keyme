$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$startVbs = Join-Path $PSScriptRoot "launch-hidden.vbs"
$stopVbs = Join-Path $PSScriptRoot "stop-hidden.vbs"
$controlVbs = Join-Path $PSScriptRoot "control-hidden.vbs"

Push-Location $root
try {
    cargo build --release
}
finally {
    Pop-Location
}

$shell = New-Object -ComObject WScript.Shell
$desktop = [Environment]::GetFolderPath("Desktop")
$startup = [Environment]::GetFolderPath("Startup")
$programs = [Environment]::GetFolderPath("Programs")
$startMenuFolder = Join-Path $programs "Keyme"
$legacyStartMenuFolder = Join-Path $programs "Keeby Windows"
New-Item -ItemType Directory -Path $startMenuFolder -Force | Out-Null

Remove-Item (Join-Path $desktop "Start Keeby Windows.lnk") -ErrorAction SilentlyContinue
Remove-Item (Join-Path $desktop "Stop Keeby Windows.lnk") -ErrorAction SilentlyContinue
Remove-Item (Join-Path $startup "Keeby Windows.lnk") -ErrorAction SilentlyContinue
Remove-Item $legacyStartMenuFolder -Recurse -ErrorAction SilentlyContinue

function New-VbsShortcut {
    param(
        [string]$Path,
        [string]$VbsPath,
        [string]$Description
    )

    $shortcut = $shell.CreateShortcut($Path)
    $shortcut.TargetPath = "$env:WINDIR\System32\wscript.exe"
    $shortcut.Arguments = "`"$VbsPath`""
    $shortcut.WorkingDirectory = $root
    $shortcut.Description = $Description
    $shortcut.Save()
}

New-VbsShortcut `
    -Path (Join-Path $desktop "Keyme.lnk") `
    -VbsPath $controlVbs `
    -Description "Open Keyme settings"

New-VbsShortcut `
    -Path (Join-Path $desktop "Start Keyme.lnk") `
    -VbsPath $startVbs `
    -Description "Start mechanical keyboard sounds"

New-VbsShortcut `
    -Path (Join-Path $desktop "Stop Keyme.lnk") `
    -VbsPath $stopVbs `
    -Description "Stop mechanical keyboard sounds"

New-VbsShortcut `
    -Path (Join-Path $startup "Keyme.lnk") `
    -VbsPath $startVbs `
    -Description "Start mechanical keyboard sounds at login"

New-VbsShortcut `
    -Path (Join-Path $startMenuFolder "Keyme.lnk") `
    -VbsPath $controlVbs `
    -Description "Open Keyme settings"

New-VbsShortcut `
    -Path (Join-Path $startMenuFolder "Start Keyme.lnk") `
    -VbsPath $startVbs `
    -Description "Start mechanical keyboard sounds"

New-VbsShortcut `
    -Path (Join-Path $startMenuFolder "Stop Keyme.lnk") `
    -VbsPath $stopVbs `
    -Description "Stop mechanical keyboard sounds"

Write-Host "Installed Keyme."
Write-Host "Desktop shortcuts: Keyme, Start Keyme, Stop Keyme"
Write-Host "Startup shortcut: Keyme"
