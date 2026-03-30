Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
quote = Chr(34)
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = scriptDir & "\Check-GeekMagicVersion.ps1"
powershell = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")
command = quote & powershell & quote & " -NoProfile -ExecutionPolicy Bypass -File " & quote & scriptPath & quote
shell.Run command, 0, True