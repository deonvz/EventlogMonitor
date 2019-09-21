'VBScript
'Purpose of script to query Errors in the eventlog
'Scripted by Deon van Zyl

Const CONVERT_TO_LOCAL_TIME = True
Dim WshShell 'Shell Commander

Set WshShell = CreateObject("WScript.Shell")
Set dtmStartDate = CreateObject("WbemScripting.SWbemDateTime")
Set dtmEndDate = CreateObject("WbemScripting.SWbemDateTime")

'DateToCheck = CDate("01/09/2005")
DateToCheck = CDate(date-1)
dtmStartDate.SetVarDate DateToCheck, CONVERT_TO_LOCAL_TIME
dtmEndDate.SetVarDate DateToCheck + 1, CONVERT_TO_LOCAL_TIME

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colLoggedEvents = objWMIService.ExecQuery _
("Select * from Win32_NTLogEvent Where Logfile = 'Application' or Logfile = 'System' and " _
& "Type = 'warnings' or Type = 'errors' and TimeWritten >= '" _ 
& dtmStartDate & "' and TimeWritten < '" & dtmEndDate & "'") 

if colLoggedEvents.Count > 0 then


	'-----------Create msg.txt (body of email-----)

	Dim objFileSystem, objOutputFile
	Dim strOutputFile

	' generate a filename base on the script name
	strOutputFile = "msg.txt"

	Set objFileSystem = CreateObject("Scripting.fileSystemObject")
	Set objOutputFile = objFileSystem.CreateTextFile(strOutputFile, TRUE)

	objOutputFile.WriteLine "Errors : " & colLoggedEvents.Count
	objOutputFile.WriteLine "==============="
	objOutputFile.WriteLine ""

	For each objEvent in colLoggedEvents
   		objOutputFile.WriteLine "Category: " & objEvent.Category
 		objOutputFile.WriteLine "Computer Name: " & objEvent.ComputerName
    		objOutputFile.WriteLine "Event Code: " & objEvent.EventCode
    		objOutputFile.WriteLine "Message: " & objEvent.Message
    		objOutputFile.WriteLine "Record Number: " & objEvent.RecordNumber
    		objOutputFile.WriteLine "Source Name: " & objEvent.SourceName
    		objOutputFile.WriteLine "Time Written: " & objEvent.TimeWritten
    		objOutputFile.WriteLine "Event Type: " & objEvent.Type
    		objOutputFile.WriteLine "User: " & objEvent.User
    		objOutputFile.WriteLine objEvent.LogFile
	Next

	objOutputFile.Close

	Set objFileSystem = Nothing

		'-----------------------------

	WshShell.Run "mailsend.bat"  'Errors in Eventlog !! Will send a email notification
end if
