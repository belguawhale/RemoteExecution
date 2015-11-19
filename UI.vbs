Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = Wscript.CreateObject("Wscript.Shell")
'Command Format: <command>`<parameter>

Function Timestamp ()	
	strSafeDate = DatePart("yyyy",Date) & "-" & Right("0" & DatePart("m",Date), 2) & "-" & Right("0" & DatePart("d",Date), 2)
	strSafeTime = Right("0" & Hour(Now), 2) & " " & Right("0" & Minute(Now), 2) & " " & Right("0" & Second(Now), 2)
	Timestamp = strSafeDate & "," & strSafeTime
End Function

Sub AppendToLog (text)
	Set WriteFile = objFSO.OpenTextFile ("Logs\log.txt", 8, False)
	WriteFile.Write Timestamp & " INPUT " & text & vbCrLf
	WriteFile.Close
End Sub

Dim Command, Reply, Temp
Do
	Reply = ""
	Command = InputBox ("Enter the command you want to run, or leave this blank to exit the UI. Command format is <command>`<parameter(s)>", "RemoteExecution UI")

	If objFSO.FileExists ("Output\output.txt") then
		objFSO.DeleteFile ("Output\output.txt")
	End if
	'In case replied too late

	if not Command = "" then
		objFSO.CreateTextFile "Input\command.txt", True
		Set objWriteMessage = objFSO.OpenTextFile ("Input\command.txt", 2)
		objWriteMessage.Write Command
		objWriteMessage.Close
		AppendToLog "Command " & Command & " sent successfully."
		WshShell.Popup "Command " & Command & " sent successfully.", 1, "Input"
		Temp = 1
		Do
			Wscript.Sleep 1000
			if objFSO.FileExists ("Output\output.txt") then
				Set objReadFile = objFSO.OpenTextFile ("Output\output.txt", 1)
				Reply = objReadFile.ReadAll
				objReadFile.Close
				objFSO.DeleteFile ("Output\output.txt")
				Temp = 0
				MsgBox Reply,, "Output"
			End If
			Temp = Temp + 1
		Loop Until Temp = 15 or Temp = 1
		If Temp = 15 then
			MsgBox "No reply received within 15000 milliseconds.",, "Output"
		End if
	End If
Loop While not Command = ""