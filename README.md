# RemoteExecution
Allows a user to remotely control a computer through Dropbox's syncing of files. Written in scripting languages for fast editing and addition of features!

If you are installing RemoteExecution, please read INSTALL located in the root of this repository.

Notable Files

.\watchdog.bat
 This is your launcher and will keep your script alive if it dies. It logs startup and restarting of the script, and will notify you of any errors.
 
.\remoterun.vbs
 This does the grunt of the work. It sits in the background, running under .\watchdog.bat and listens for any commands. You may edit this as you please, but remember to send the command `exit\`` to kill the old copy and execute the new script.
 
.\UI.vbs
 This is the preferred user interface from a computer, since it logs any input, and receives output. Very useful for checking if the script is running or not using `ping\``, since it returns a reply. I personally use IFTTT for sending commands over phone; just configure IFTTT to add a text file "command.txt" under Dropbox\RemoteExecution\Input.

Logs of startup/restarting of the script are found under .\Logs\output.txt, and logs of all commands (INPUT and OUTPUT) are found under .\Logs\log.txt.
For help, simply send the command `help\``.

Notes

- Changes to remoterun.vbs will not take effect until you send `exit\`` to restart it. If the script is not responding to anything afterwards, there may be a syntax error in the script. You can troubleshoot that by seeing how often the script is restarting in output.txt. In that case, fix the script and save it. Watchdog.bat will constantly restart the script, and so your changes will take effect.
- restarteverything.vbs terminates watchdog.bat in case you want to make changes to it.
- Screenshots are very useful for viewing what your commands did!
- Don't send commands faster than 1 per 15 seconds through text message, since it may create extra command (#).txt files that will never get executed.
- Don't forget the "\`" separator between commands and their parameters! (That includes commands with no parameters; send `<command>\`` to execute them)
