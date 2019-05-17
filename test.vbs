Dim mytime,myhour,myminutes,mysecondes,heureapparition,temp_attente,temp_restant,mysec,Seconde_Apparition

set WshShell = WScript.CreateObject("WScript.Shell")
Message = InputBox("Entrez votre message","Minuterie")
Seconde_voulu=InputBox("Combien de seconde(s) ? ","Minuterie")
mytime = Now
myhour= Hour(mytime)
mymin=Minute(mytime)
mysec=Second(mytime)
check=0
temp_attente=Seconde_voulu

while check=0
	if temp_attente=0 then 
		MsgBox Message
		check=1
	else
		WshShell.Popup "Il reste : " & temp_attente & " secondes",1,"title"
		temp_attente=temp_attente-1
	end if
wend 
'intButton = object.Popup(strText,[nSecondsToWait],[strTitle],[nType])
