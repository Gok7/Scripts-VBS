Dim fso, f, f1, s, sf
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFolder("C:\Users\john\Music")
   Set sf = f.SubFolders
   For Each f1 in sf
      s = s & f1.name 
      s = s & vbCrLf
   Next

   MsgBox(s)