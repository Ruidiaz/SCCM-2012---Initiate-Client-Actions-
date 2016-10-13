'------------------------------------------------------------------------------------
'	InitiateSCCM.vbs
' 		- Script to invoke ConfigMgr client actions
'
'	Version	Date		   Who				        Changes
'	1.0		13/10/2016   Javier RUIDIAZ   	.
'------------------------------------------------------------------------------------

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly = &h20
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

On Error Resume Next
Dim objShell, objFSO
Dim WriteLogfile, strlogFile

strComputer = "."

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject") 


' log file settings
WriteLogfile = 1	' change to '0' if no logging is required
LogfilePath = "C:\IT_Logs"	

'=== Logfile Path and Name ; create ===
strTimeStamp = fnGetTimeStamp()
strLogfileName = "InitiateSCCM_" & strTimeStamp & ".log"	'amend logfile name
strLogfile = LogfilePath & "\" & strLogfileName

if WriteLogfile = 1 Then
	If Not objFSO.FolderExists(LogfilePath) Then
		Set createLogFolder = objFSO.CreateFolder(LogfilePath)
	end if
end if


' main part
writeLog("--------------------------------------------------------------------------------------------------")
writeLog("INFO: InitiateSCCM started " & strTimeStamp & ".")
writeLog("--------------------------------------------------------------------------------------------------")

	'Set Variables For Actions 
	actionNameToRun = "Updates Source Scan Cycle" 
	actionNameToRun1 = "Software Updates Assignments Evaluation Cycle" 
	
	'Create and use the control panel applet for client actions 
	Dim controlPanelAppletManager 
	Set controlPanelAppletManager = CreateObject("CPApplet.CPAppletMgr") 
	Dim clientActions 
	Set clientActions = controlPanelAppletManager.GetClientActions() 
	Dim clientAction 
	
	'Find which actions are available 
	For Each clientAction In clientActions 
	
	'List available client actions, output using the Name property (below). 
			writeLog("Action: " & clientAction.Name )
			
	'Run statements per results 
			If clientAction.Name = actionNameToRun Then 
					clientAction.PerformAction 
					 
			End If 
			 
			If clientAction.Name = actionNameToRun1 Then 
					clientAction.PerformAction 
					 
			End If 
	Next
	WScript.Sleep(180000) 
	writeLog("Executed: " & actionNameToRun )
	writeLog("Executed: " & actionNameToRun1 )


writeLog("--------------------------------------------------------------------------------------------------")
writeLog("INFO: InitiateSCCM has finished")
writeLog("--------------------------------------------------------------------------------------------------")


' functions and subs
'=============================================
'	Function for logging
'=============================================
Function writeLog(strLogging)
	if WriteLogfile = 1 Then
		Set objLog = objFSO.OpenTextFile(strLogfile, ForAppending, True)
		objlog.writeline Date & " " & Time & " " & " <" & strLogging & ">"
		objLog.close
	end if
End Function

'=============================================
'	Functions for formatting MsgBoxes
'=============================================
Function MsgBoxInfo(strTXT)
	MsgBox strTXT, vbInformation, "INFO"
End Function

Function MsgBoxExcl(strTXT)
	MsgBox strTXT, vbExclamation, "ATTENTION"
End Function

Function MsgBoxErr(strTXT)
	MsgBox strTXT, vbCritical, "ERROR"
End Function

'=============================================
'	Function for formatting TimeStamps
'=============================================
Function fnGetTimeStamp()
	t = Now
	
	y = Year(t)
	m = Month(t)
	d = Day(t)
	h = Hour(t)
	mn = Minute(t)
	s = Second(t) 
	
	strArr = Array(m,d,h,mn,s)
	
	for i=0 To 4
		if strArr(i) < 10 Then
			strArr(i) = "0" & strArr(i)
		End If
	Next
	
	fnGetTimeStamp = y & strArr(0) & strArr(1) & "_" & strArr(2) & strArr(3) & strArr(4)
End Function
