Dim objArgs, I
Set objArgs = WScript.Arguments
For I = 0 to objArgs.Count -1
	Message = Message & "param�tre" & I & " = " & objArgs(I) & VbCrlf
Next
MsgBox Message
