Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c nuclide.bat"
oShell.Run strArgs, 0, false