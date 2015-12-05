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
