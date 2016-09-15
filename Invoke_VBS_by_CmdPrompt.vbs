set oShell = WScript.CreateObject("WScript.shell")
oShell.Run "cmd"
WScript.Sleep 100
oShell.sendkeys "cd /"
oShell.sendkeys "{ENTER}"
oShell.sendkeys "c:\Windows\SysWOW64\wscript.exe ""C:\Users\bsubrama091609\Desktop\Launching_By_TestSet.vbs"""
oShell.sendkeys "{ENTER}"
