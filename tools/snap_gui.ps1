# Запуск окна программы, снимок окна в tools\snap.jpg.b64 и закрытие (проверка интерфейса)
param([string]$Root, [string]$Script = "app\gui_app.py")
Set-Location $Root
Remove-Item tools\hwnd.txt -ErrorAction SilentlyContinue
$p = Start-Process C:\Python314\python.exe -ArgumentList $Script -PassThru -RedirectStandardError "$env:TEMP\gui_err.txt"
Start-Sleep 6; $p.Refresh()
$h = $p.MainWindowHandle
if (Test-Path tools\hwnd.txt) { $h = [long](Get-Content tools\hwnd.txt); Remove-Item tools\hwnd.txt }
& powershell -ExecutionPolicy Bypass -File tools\snap_window.ps1 -Handle $h -Out "$env:TEMP\snap.png"
Stop-Process -Id $p.Id
Get-Content "$env:TEMP\gui_err.txt"
[Convert]::ToBase64String([IO.File]::ReadAllBytes("$env:TEMP\snap.png")) | Set-Content "tools\snap.b64" -Encoding ascii
