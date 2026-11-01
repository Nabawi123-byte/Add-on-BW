' Menjalankan server Master di latar belakang (tanpa jendela).
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
sh.Run "cmd /c node server.js > master-server.log 2>&1", 0, False
