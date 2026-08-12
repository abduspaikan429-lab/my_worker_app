Set ws = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strPath = fso.GetParentFolderName(WScript.ScriptFullName)
ws.CurrentDirectory = strPath
ws.Run "cmd /c """ & strPath & "\启动系统.bat""", 0
