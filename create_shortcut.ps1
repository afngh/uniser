$PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop "PIN Lookup.lnk"

# Check if PyInstaller dist\PinLookup.exe exists
$exePath = Join-Path $PSScriptRoot "dist\PinLookup.exe"
if (Test-Path $exePath) {
    $targetPath = (Resolve-Path $exePath).Path
    $targetArgs = ""
} else {
    $pythonwCmd = Get-Command pythonw -ErrorAction SilentlyContinue
    if ($pythonwCmd) {
        $targetPath = $pythonwCmd.Source
    } else {
        $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
        if ($pythonCmd) {
            $targetPath = $pythonCmd.Source -replace "python\.exe$", "pythonw.exe"
        } else {
            $targetPath = "pythonw.exe"
        }
    }
    $appPyPath = Join-Path $PSScriptRoot "app.py"
    $targetArgs = "`"$appPyPath`""
}

$WshShell = New-Object -ComObject WScript.Shell
$shortcut = $WshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $targetPath
if ($targetArgs) {
    $shortcut.Arguments = $targetArgs
}
$shortcut.WorkingDirectory = $PSScriptRoot
$shortcut.IconLocation = "$env:SystemRoot\System32\shell32.dll,13"
$shortcut.Save()

Write-Host "Desktop shortcut created successfully!" -ForegroundColor Green
Write-Host "Location: $shortcutPath"
Write-Host "Target:   $targetPath $targetArgs"
