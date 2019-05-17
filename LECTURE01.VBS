'afficher l'espace libre, la taille totale et le nom de tous les lecteurs

Function ShowDriveList
   Dim fso, d, dc, s, n, disk1
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set dc = fso.Drives
   For Each d in dc
      n = ""
      s = s & d.DriveLetter & ":\   " 
      If d.DriveType = Remote Then
         n = d.ShareName & FormatNumber(d.FreeSpace/1073741824, 0) & " GB " & "Free of " & FormatNumber(d.TotalSize/1073741824, 0) & " GB "
      ElseIf d.IsReady Then
         n = d.VolumeName & FormatNumber(d.FreeSpace/1073741824, 0) & " GB " & "Free of " & FormatNumber(d.TotalSize/1073741824, 0) & " GB "
      End If
      s = s & n & VbCrlf
   Next
   ShowDriveList = s
End Function

MsgBox(ShowDriveList)
