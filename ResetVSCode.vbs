Dim version, user, vscode, code, history, folder

version = "1.108.2"

Set shell = WScript.CreateObject("WScript.Shell")

user = shell.ExpandEnvironmentStrings("%USERPROFILE%")

Set fso = CreateObject("Scripting.FileSystemObject")

vscode = user & "\.vscode"

If fso.FolderExists(vscode) Then
    fso.DeleteFolder vscode, True
End If

code = user & "\AppData\Roaming\Code"

If fso.FolderExists(code) Then
    fso.DeleteFolder code, True
End If

history = user & _
          "\AppData" & _
          "\Roaming" & _
          "\Microsoft" & _
          "\Windows" & _
          "\PowerShell" & _
          "\PSReadLine" & _
          "\ConsoleHost_history.txt"

If fso.FileExists(history) Then
    fso.DeleteFile history, True
End If

shell.Run "D:\vscode" & _
          "\VSCode-win32-x64-" & version & _
          "\Code.exe", 1, False

Set shell = Nothing

folder = user & _
         "\AppData" & _
         "\Roaming" & _
         "\Code" & _
         "\User"

If Not fso.FolderExists(folder) Then
    CreateNestedFolder fso, folder
End If

Set f = fso.CreateTextFile(folder & "\settings.json", True)
f.WriteLine "{"
f.WriteLine "    ""chat.disableAIFeatures"": true,"
f.WriteLine "    ""workbench.tree.indent"": 18,"
f.WriteLine "    ""workbench.startupEditor"": ""none"","
f.WriteLine "    ""workbench.secondarySideBar.defaultVisibility"": ""hidden"""
f.Write "}"
f.Close

Set fso = Nothing

Sub CreateNestedFolder(fso, path)
    If Not fso.FolderExists(path) Then
        CreateNestedFolder fso, fso.GetParentFolderName(path)
        fso.CreateFolder(path)
    End If
End Sub
