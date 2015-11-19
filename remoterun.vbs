Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell=Wscript.CreateObject("Wscript.Shell")
'Command format: <command>`<parameter>
Dim Command, Temp, HelpMsg

Function Timestamp ()	
	strSafeDate = DatePart("yyyy",Date) & "-" & Right("0" & DatePart("m",Date), 2) & "-" & Right("0" & DatePart("d",Date), 2)
	strSafeTime = Right("0" & Hour(Now), 2) & " " & Right("0" & Minute(Now), 2) & " " & Right("0" & Second(Now), 2)
	Timestamp = strSafeDate & "," & strSafeTime
End Function

Sub AppendToLog (text)
	Wscript.Sleep 500
	Set WriteFile = objFSO.OpenTextFile ("Logs\log.txt", 8, False)
	WriteFile.Write Timestamp & " OUTPUT " & text & vbCrLf
	WriteFile.Close
	Wscript.Echo Timestamp & " OUTPUT " & text
End Sub

Sub WriteReply (text)
	objFSO.CreateTextFile "Output\output.txt", True
	Set WriteFile = objFSO.OpenTextFile ("Output\output.txt", 2)
	WriteFile.Write text
	WriteFile.Close
	AppendToLog "Sent reply " & text
End Sub

Sub SendHelp (command)
	If command = "" then
		Set ReadHelp = objFSO.OpenTextFile ("Help\help.txt", 1)
		HelpMsg = ReadHelp.ReadAll
		ReadHelp.Close
		objFSO.CreateTextFile "Output\output.txt", True
		Set WriteFile = objFSO.OpenTextFile ("Output\output.txt", 2)
		WriteFile.Write HelpMsg
		WriteFile.Close
		AppendToLog "Sent general help"
	Elseif objFSO.FileExists ("Help\" & command & ".txt") then
		Set ReadHelp = objFSO.OpenTextFile ("Help\" & command & ".txt", 1)
		HelpMsg = ReadHelp.ReadAll
		ReadHelp.Close
		objFSO.CreateTextFile "Output\output.txt", True
		Set WriteFile = objFSO.OpenTextFile ("Output\output.txt", 2)
		WriteFile.Write HelpMsg
		WriteFile.Close
		AppendToLog "Sent help for command " & command
	Else
		WriteReply "No help for " & command & " exists."
	End If
End Sub

'This is so it doesn't get into a crash loop
If objFSO.FileExists ("Input\command.txt") then
	objFSO.DeleteFile ("Input\command.txt")
End if

do
	Wscript.Sleep 2000
	If objFSO.FileExists ("Input\command.txt") then
		Set ReadFile = objFSO.OpenTextFile ("Input\command.txt", 1)
		Command = ReadFile.ReadLine
		ReadFile.Close
		objFSO.DeleteFile ("Input\command.txt")
		Wscript.Sleep 2000
		'wait for dropbox to sync input
		AppendToLog "Received Input: " & Command
		if InStr (Command, "`") then
			Temp = Split (Command, "`")
			if Temp(0) = "run" then
				WshShell.Run Temp(1)
				WriteReply "Successfully executed " & Temp(1)
			elseif Temp(0) = "nircmd" then
				WshShell.Run "Utilities\nircmd.exe " & Temp(1)
				WriteReply "Successfully executed nircmd.exe " & Temp(1)
			elseif Temp(0) = "nircmdc" then
				WshShell.Run "Utilities\nircmdc.exe " & Temp(1)
				WriteReply "Successfully executed nircmdc.exe " & Temp(1)
			elseif Temp (0) = "takescreenshot" then
				if Temp(1) = "" then
					WshShell.Run ("Utilities\nircmd.exe savescreenshot " & chr(34) & "Screenshots\" & Timestamp & ".png" & chr(34))
					WriteReply "Successfully took screenshot with name " & Timestamp & ".png"
				else
					WshShell.Run ("Utilities\nircmd.exe savescreenshot " & chr(34) & "Screenshots\" & Temp(1) & ".png" & chr(34))
					WriteReply "Successfully took screenshot at " & Timestamp & " with name " & Temp(1) & ".png"
				end if
			elseif Temp (0) = "sendkeys" then
				if not Temp(1) = "" then
					WshShell.SendKeys Temp (1)
					WriteReply "Successfully sent keystrokes " & Temp (1)
				end if
			elseif Temp(0) = "log" then
				AppendToLog Temp(1)
			elseif Temp(0) = "reply" then
				WriteReply Temp(1)
			elseif Temp(0) = "ping" then
				WriteReply "Pong!"
			elseif Temp(0) = "moo" then
				WriteReply "MoooOoOOoOOooOOOooOoOOoooOOooooOOOooo!"
			elseif Temp(0) = "help" then
				SendHelp Temp(1)
			elseif Temp(0) = "exit" then
				if Temp(1) = "" then
					WriteReply "Successfully quit at " & Timestamp
					Wscript.Quit
				else
					WriteReply "Successfully quit with reason: " & Temp(1)
					Wscript.Quit
				end if
			else
				WriteReply "Command invalid. Send help`commands for a listing of commands."
			end if
		Else
			WriteReply "Command missing a `. If you don't want to specify parameters, just send command`."
		End If
	End If
loop