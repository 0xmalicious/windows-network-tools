'*********************************************************
'* Les partages r�seau                                   *
'* DOCm�mo - www.docmemo.com                             *
'*********************************************************

strComputer = "."
Liste = "Les partages r�seau" & vbCR
Liste = Liste & "---------------------------" & vbCR & vbCR
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")
For each objShare in colShares 
	Liste = Liste & "Nom : " & objShare.Name
	Liste = Liste & " - L�gende : " & objShare.Caption & vbCR
	Liste = Liste & "Chemin : " & objShare.Path & vbCR
	Liste = Liste & vbCR
Next
MsgBox Liste,,"Les ressources partag�es"
