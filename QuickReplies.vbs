Set shell = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
script = fs.BuildPath(fs.GetParentFolderName(WScript.ScriptFullName), "app.ps1")
ps = shell.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
shell.Run Chr(34) & ps & Chr(34) & " -NoProfile -STA -ExecutionPolicy Bypass -File " & Chr(34) & script & Chr(34), 0, True
