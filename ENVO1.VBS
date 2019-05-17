Dim objArgs, I
Set WShell = WScript.CreateObject("WScript.Shell")
Set MesVariablesEnvironnement = WShell.environment

for each variable in MesVariablesEnvironnement
	message = message & variable & VbCrlf
Next
MsgBox Message