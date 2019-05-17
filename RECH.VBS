Dim Montagnes,input
Set Montagnes=CreateObject("Scripting.Dictionary")
Montagnes.Add "1500","Mont Aigoual"
Montagnes.Add "4800","Mont Blanc"
Montagnes.Add "2800","Pic du Midi"

do until input="fin"
	input=InputBox("Quelle est l'hauteur ?","Question","")
   If Montagnes.Exists(input) Then
         msg = "à " & input & "m on trouve le " & Montagnes.Item(input)
   Else
      msg = "Aucune entrée"
   End If
	MsgBox msg
Loop


