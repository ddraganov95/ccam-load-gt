Set shell = CreateObject("WScript.Shell")

shell.Run """C:\Program Files\PowerShell\7\pwsh.exe"" -NoProfile -ExecutionPolicy Bypass -File ""D:\Users\ddraganov\ccam-load-gt\check-ccam.ps1""", 0, True