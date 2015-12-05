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
'==============================================================
