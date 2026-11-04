Set shell = CreateObject("WScript.Shell")

shell.Run """C:\Program Files\PowerShell\7\pwsh.exe"" -NoProfile -ExecutionPolicy Bypass -File "".\check-ccam.ps1""", 0, True