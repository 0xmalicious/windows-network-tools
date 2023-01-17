'*********************************************************
'* Les partages réseau                                   *
'* DOCmémo - www.docmemo.com                             *
'*********************************************************

strComputer = "."
Liste = "Les partages réseau" & vbCR
Liste = Liste & "---------------------------" & vbCR & vbCR
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")
For each objShare in colShares 
	Liste = Liste & "Nom : " & objShare.Name
	Liste = Liste & " - Légende : " & objShare.Caption & vbCR
	Liste = Liste & "Chemin : " & objShare.Path & vbCR
	Liste = Liste & vbCR
Next
MsgBox Liste,,"Les ressources partagées"
