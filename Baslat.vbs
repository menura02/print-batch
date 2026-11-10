Option Explicit
Dim shell, fso, base, candidates, candidate, python, localApp, command
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(WScript.ScriptFullName)
localApp = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
candidates = Array(base & "\runtime\pythonw.exe", shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe", localApp & "\Programs\Python\Python313\pythonw.exe", localApp & "\Programs\Python\Python312\pythonw.exe")
python = ""
For Each candidate In candidates
    If fso.FileExists(candidate) Then
        python = candidate
        Exit For
    End If
Next
If python = "" Then
    MsgBox "Python 3.10 veya ustu bulunamadi. Python kurun veya python printer_app.py komutunu kullanin.", 16, "PrintBatch"
Else
    shell.CurrentDirectory = base
    command = Chr(34) & python & Chr(34) & " " & Chr(34) & base & "\printer_app.py" & Chr(34)
    shell.Run command, 0, False
End If
