Dim objArgs, I
Set objArgs = WScript.Arguments
For I = 0 to objArgs.Count -1
	Message = Message & "paramètre" & I & " = " & objArgs(I) & VbCrlf
Next
MsgBox Message
