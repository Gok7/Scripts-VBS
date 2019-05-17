Dim WshShell,strDesktop
Set WShell = WScript.CreateObject("WScript.Shell")
strDesktop= WShell.SpecialFolders("Desktop")

set ShellLink = WShell.CreateShortcut(strDesktop & "\Bloc nokkkte de windows.lnk")
ShellLink.TargetPath  = "C:\Windows\system32\notepad.exe"
ShellLink.Save

set ShellLink = WShell.CreateShortcut(strDesktop & "\Calculatrice.lnk")
ShellLink.TargetPath  = "C:\Windows\system32\calc.exe"
ShellLink.Save