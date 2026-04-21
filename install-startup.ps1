$ErrorActionPreference = 'Stop'

$startupFolder = [System.Environment]::GetFolderPath('Startup')
$shortcutPath = Join-Path $startupFolder 'Match Dashboard.lnk'
$targetVbs = Join-Path $PSScriptRoot 'start.vbs'

$wshShell = New-Object -ComObject WScript.Shell
$shortcut = $wshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = 'wscript.exe'
$shortcut.Arguments = '"' + $targetVbs + '"'
$shortcut.WorkingDirectory = $PSScriptRoot
$shortcut.WindowStyle = 7  # Minimized
$shortcut.Description = 'Match Dashboard auto-start'
$shortcut.Save()

[System.Windows.Forms.MessageBox]::Show("등록 완료!`n`nPC 부팅 시 자동으로 실행됩니다.", "Match Dashboard", "OK", "Information") | Out-Null
