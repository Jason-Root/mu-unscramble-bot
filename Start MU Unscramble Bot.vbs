Option Explicit

Dim fso, shellApp, scriptDir, pythonwPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set shellApp = CreateObject("Shell.Application")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonwPath = scriptDir & "\.venv\Scripts\pythonw.exe"

If Not fso.FileExists(pythonwPath) Then
    MsgBox "pythonw.exe was not found at:" & vbCrLf & pythonwPath, vbCritical, "MU Unscramble Bot"
    WScript.Quit 1
End If

shellApp.ShellExecute pythonwPath, "-m mu_unscramble_bot", scriptDir, "runas", 1
