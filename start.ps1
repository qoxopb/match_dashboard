$ErrorActionPreference = 'Continue'
Set-Location $PSScriptRoot

# Win32 API for minimizing Chrome
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32Api {
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@ -ErrorAction SilentlyContinue

function Test-Port($port) {
    try {
        Get-NetTCPConnection -LocalPort $port -State Listen -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Minimize-Chrome {
    Start-Sleep -Milliseconds 800
    Get-Process chrome -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.MainWindowHandle -ne 0) {
            [Win32Api]::ShowWindow($_.MainWindowHandle, 6) | Out-Null  # 6 = SW_MINIMIZE
        }
    }
}

# 0) npm install (node_modules 없으면)
if (-not (Test-Path "$PSScriptRoot\node_modules")) {
    Write-Host "node_modules 없음 → npm install 실행 중..."
    Push-Location $PSScriptRoot
    npm install
    Pop-Location
}

# 1) CDP Chrome
if (-not (Test-Port 9222)) {
    Start-Process "C:\Program Files\Google\Chrome\Application\chrome.exe" -ArgumentList '--remote-debugging-port=9222', '--user-data-dir=C:\chrome-debug', '--start-minimized' -WindowStyle Minimized
    Start-Sleep -Seconds 2
    Minimize-Chrome
}

# 2) Node server — 기존 실행 중이면 종료 후 새로 실행
$nodeWasRunning = $false
$existing = Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue
if ($existing) {
    $nodeWasRunning = $true
    try {
        Stop-Process -Id $existing.OwningProcess -Force -ErrorAction SilentlyContinue
    } catch {}
    # 포트 해제 대기
    $waited = 0
    while ((Test-Port 3000) -and $waited -lt 5) {
        Start-Sleep -Milliseconds 500
        $waited++
    }
}

$pinfo = New-Object System.Diagnostics.ProcessStartInfo
$pinfo.FileName = "node.exe"
$pinfo.Arguments = "server.js"
$pinfo.WorkingDirectory = $PSScriptRoot
$pinfo.CreateNoWindow = $true
$pinfo.UseShellExecute = $false
$pinfo.WindowStyle = 'Hidden'
[System.Diagnostics.Process]::Start($pinfo) | Out-Null

# 서버 포트 열릴 때까지 대기
$waited = 0
while (-not (Test-Port 3000) -and $waited -lt 15) {
    Start-Sleep -Seconds 1
    $waited++
}

# 3) 브라우저 열기 (기존 노드가 떠있던 재시작 케이스에는 생략)
if (-not $nodeWasRunning) {
    Start-Process "http://localhost:3000"
}

# 4) 토스트 알림 (등록된 PowerShell AppID 사용)
try {
    [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType=WindowsRuntime]
    [void][Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications, ContentType=WindowsRuntime]
    [void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType=WindowsRuntime]

    $xmlString = @"
<toast>
  <visual>
    <binding template="ToastGeneric">
      <text>Match Dashboard</text>
      <text>이제 사용 가능합니다. (최소화된 크롬 브라우저를 닫지 마세요)</text>
    </binding>
  </visual>
</toast>
"@

    $xml = New-Object Windows.Data.Xml.Dom.XmlDocument
    $xml.LoadXml($xmlString)
    $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
    $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show($toast)
} catch {
    # 폴백: Windows Forms 풍선 알림
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        $notify = New-Object System.Windows.Forms.NotifyIcon
        $notify.Icon = [System.Drawing.SystemIcons]::Information
        $notify.BalloonTipTitle = 'Match Dashboard'
        $notify.BalloonTipText = "이제 사용 가능합니다. (최소화된 크롬 브라우저를 닫지 마세요)"
        $notify.Visible = $true
        $notify.ShowBalloonTip(10000)
        Start-Sleep -Seconds 8
        $notify.Dispose()
    } catch {}
}
