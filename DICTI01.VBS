Dim Montagnes, TableauElements,TableauCles,Boucle,Message
Set Montagnes=CreateObject("Scripting.Dictionary")
Montagnes.Add 1500,"Mont Aigoual"
Montagnes.Add 4800,"Mont Blanc"
Montagnes.Add 2800,"Pic du Midi"

MsgBox"A 4800m on trouve le :"& Montagnes.Item(4800)

TableauElements=Montagnes.Items
TableauCles=Montagnes.keys

For boucle=0 to Montagnes.count - 1
Message = Message & TableauElements(boucle) & " " & TableauCles(boucle) & VbCrlf
Next
MsgBox(Message)

