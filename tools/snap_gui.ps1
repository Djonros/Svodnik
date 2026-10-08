# Запуск окна программы, снимок окна в tools\snap.jpg.b64 и закрытие (проверка интерфейса)
param([string]$Root, [string]$Script = "app\gui_app.py")
Set-Location $Root
$p = Start-Process C:\Python314\python.exe -ArgumentList $Script -PassThru -RedirectStandardError "$env:TEMP\gui_err.txt"
Start-Sleep 6; $p.Refresh()
& powershell -ExecutionPolicy Bypass -File tools\snap_window.ps1 -Handle $p.MainWindowHandle -Out "$env:TEMP\snap.png"
Stop-Process -Id $p.Id
Get-Content "$env:TEMP\gui_err.txt"
[Convert]::ToBase64String([IO.File]::ReadAllBytes("$env:TEMP\snap.png")) | Set-Content "tools\snap.b64" -Encoding ascii
