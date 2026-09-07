' Runs Get-ComputerVersions on Silent Mode
' 0 = hide window, True = wait until finished
Set objShell = CreateObject("WScript.Shell")
objShell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -File ""\\irt-fss\irt-fs2\SERVICE\IT-Hiwi\SWVV-Modern\Daten\Skript\Get-ComputerVersions.ps1""", 0, True
