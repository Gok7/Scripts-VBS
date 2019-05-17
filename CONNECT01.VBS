Dim WshNetwork
Set WshNetwork = WScript.CreateObject("WScript.Network")
WshNetwork.MapNetworkDrive "z:", "\\Server\Public"