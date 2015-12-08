'###
SAvblVersionName = "PreMerge"
SAvblVersionNumber = "0.1"

'##############################################################
'### Error detection and handling
'==============================================================
Sub ShowErr
	Wscript.Echo "-------------------" 
	Wscript.Echo "Err:"
	Wscript.Echo "Err.Number: " & Err.Number
	Wscript.Echo Err.Description
	Wscript.Echo Err.Source
	Wscript.Echo "-------------------" 
End sub

'##############################################################
'### Input
'==============================================================
function ReadFile2Array(strFileLocation) ' v.1 - Returns array with file's content, or array with "ReadFile:Fail" string at the 0 record, if fails to read the file,
	'used functions: Logger
	'used globals: booVerbose
	'------------------------------------------------------------------
	if booVerbose then wscript.echo "File2Arr>>" & strFileLocation & vbCrlf
	'------------------------------------------------------------------
	dim objFSO 		: set objFSO = CreateObject("Scripting.FileSystemObject")
	dim objFile
	dim arrFileRows()	: redim preserve arrFileRows(0)
	dim intFileRowCNT	: intFileRowCNT = 0
	'------------------------------------------------------------------
	if objFSO.FileExists(strFileLocation) then
		on error resume next
		Set objFile = objFSO.OpenTextFile(strFileLocation, 1)
		if (Err.Number = 53) then
			arrFileRows(intFileRowCNT) = "ReadFile:Fail"
			strTmpStr = "ReadFile:Fail:" & strFileLocation &" Err:"&Err.Source & "/" & Err.Number & "/" & Err.Description
			Err.Clear
			Logger "ERROR : ReadFile2Array : " & strTmpStr, wscript.ScriptFullName & ".log"
			exit function
		end if
		on error goto 0
	
		Do Until objFile.AtEndOfStream
			redim preserve arrFileRows(intFileRowCNT)
			arrFileRows(intFileRowCNT) = objFile.ReadLine
			intFileRowCNT = intFileRowCNT + 1
		Loop
		objFile.Close
	else
		arrFileRows(intFileRowCNT) = "ReadFile:Fail"
	end if
	ReadFile2Array = arrFileRows
End Function

'##############################################################
'### Logging & Log handling
'==============================================================
Sub Logger(strLog, strLogFile) '# v.1 - Logger "Log message", Logfile
	'------------------------------------------------------------------
	dim objLog
	dim objFSO : set objFSO = CreateObject("Scripting.FileSystemObject")
	'------------------------------------------------------------------
	if objFSO.fileexists(strLogFile) then 
		Set objLog = objFSO.OpenTextFile(strLogFile, 8)
	else
		Set objLog = objFSO.CreateTextFile(strLogFile)
	End If
	objLog.WriteLine(Date & " " & Time &" - " & strLog)
	objLog.Close
	set objLog = nothing
End Sub

'#####################################
'### Voice
'==============================================================
sub Speak(strTextToSpeak)
	'CreateObject("SAPI.SPvoice").Speak"text"
	CreateObject("SAPI.SPvoice").Speak strTextToSpeak
end sub
'==============================================================

'##############################################################
'### Help
'==============================================================
sub ShowHelp
	WScript.Echo "-------------------------------------------------------------------------------"
	WScript.Echo " SAvbl - v.: " & SAvblVersionNumber & " - " & """" & SAvblVersionName & """"
	WScript.Echo "-------------------------------------------------------------------------------"
	WScript.Echo
	WScript.Echo "   Usage:"
	WScript.Echo 
	WScript.Echo  
	WScript.Echo "	/LOG:<file path> - Creates log about deleted items. (overwrites existing log)" 
	WScript.Echo "	/LOG+:<file path> - Creates log about deleted items. (append to existing log)"
	WScript.Echo "	/LOG - Creates log with default log name (.date-time.log), to the script's dir.)"
	WScript.Echo 
end sub
'==============================================================

'##############################################################
'### Testing
'==============================================================
sub ShowArray(strLabel, arrToShow)
	'------------------------------------------------------------------
	dim intArrCntr 	: intArrCntr = 0
	'------------------------------------------------------------------
	wscript.echo vbCrLf & "-----[ " & strLabel & " ]---"
	for intArrCntr=0 to ubound(arrToShow)
		wscript.echo intArrCntr & "." & arrToShow(intArrCntr) & "."
	next
	wscript.echo "-------------------------------------------------------"
end sub
