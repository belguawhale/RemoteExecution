Set WshShell=Wscript.CreateObject("Wscript.Shell")
Wscript.Sleep 10000
WshShell.Run ("taskkill /im cmd.exe /f")
WshShell.Run ("watchdog.bat")