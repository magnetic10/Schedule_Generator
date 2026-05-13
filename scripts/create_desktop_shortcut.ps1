param(
    [string]$TargetPath = ""
)

$ErrorActionPreference = "Stop"

if (-not $TargetPath) {
    $Root = Split-Path -Parent $PSScriptRoot
    $TargetPath = Join-Path $Root "dist\work-scheduler-v3\work-scheduler-v3.exe"
}

$ResolvedTarget = Resolve-Path -LiteralPath $TargetPath
$Desktop = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = Join-Path $Desktop "근무표 생성기.lnk"

$Shell = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $ResolvedTarget.Path
$Shortcut.WorkingDirectory = Split-Path -Parent $ResolvedTarget.Path
$Shortcut.Description = "근무표 생성기 실행"
$Shortcut.Save()

Write-Host "Shortcut created:"
Write-Host $ShortcutPath
