' 自动定位：脚本在哪个文件夹，就扫描哪个文件夹下的 办公常用 和 开发常用
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
shell.CurrentDirectory = scriptDir
shell.Run "D:\anaconda3\python.exe """ & scriptDir & "\generate_inventory.py""", 1, True
