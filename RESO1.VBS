         Set WshNetwork = WScript.CreateObject("WScript.Network")
		 msg = "Nom de l'ordi = " & WshNetwork.ComputerName
		 msg = msg & VbCrLf & "Nom utilisateur = " & WshNetwork.UserName
		 
MsgBox msg
