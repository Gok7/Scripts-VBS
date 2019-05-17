dim Valeur_entree,Somme,Compteur
Somme=0
	
	Valeur_entree=InputBox("Entrez une valeur :","titre")
	Compteur=Valeur_entree
    Do
	Somme=Somme+Compteur
	Compteur=Compteur-1
    Loop while Compteur>0
	MsgBox "La somme est :" &Somme
   