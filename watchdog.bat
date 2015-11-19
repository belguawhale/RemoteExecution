@echo off
title RemoteExecution Client
:a
echo.>>.\Logs\output.txt
echo RemoteExecution successfully started.
echo RemoteExecution successfully started.>>.\Logs\output.txt
echo %DATE% %TIME%>>.\Logs\output.txt
cscript.exe remoterun.vbs
echo RemoteExecution terminated. Restarting in 10 seconds
echo RemoteExecution terminated. Restarting in 10 seconds...>>.\Logs\output.txt
ping -n 10 127.0.0.1>NUL
goto a