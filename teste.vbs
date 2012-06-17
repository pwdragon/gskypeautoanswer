'http://stackoverflow.com/questions/2806584/how-do-i-run-a-vbscript-in-32-bit-mode-on-a-64-bit-machine
'How do I run a VBScript in 32-bit mode on a 64-bit machine?
'* Click Start, click Run, type %windir%\SysWoW64\cmd.exe, and then click OK.
'cscript vbscriptfile.vbs

 '// Create a Skype4COM object:
 Set oSkype = WScript.CreateObject("Skype4COM.Skype", "Skype_")
 
 '// Place a call to echo123:
 Set oCall = oSkype.PlaceCall("echo123")
 
 '// Call status events:
 Public Sub Skype_CallStatus(ByRef aCall, ByVal aStatus)
   WScript.Echo ">Call " & aCall.Id & " status " & oSkype.Convert.CallStatusToText(aStatus)
 End Sub
