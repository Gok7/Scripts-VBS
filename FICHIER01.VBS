Dim fso, dossier1,dossier 

Set fso = CreateObject("Scripting.FileSystemObject")

Set dossier1 = fso.GetFolder("c:\TEMP1")
Set dossier2 = fso.GetFolder("c:\TEMP2")
dossier1.Copy (dossier2)