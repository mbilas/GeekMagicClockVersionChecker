Set shell = CreateObject("WScript.Shell")
quote = Chr(34)
command = quote & "C:\windows\System32\WindowsPowerShell\v1.0\powershell.exe" & quote & " -NoProfile -ExecutionPolicy Bypass -File " & quote & "D:\Scripts\GeekMagicClock\Check-GeekMagicVersion.ps1" & quote
shell.Run command, 0, True