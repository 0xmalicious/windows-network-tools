'*********************************************************
'* La liste des programmes installés                     *
'* DOCmémo - www.docmemo.com                             *
'*********************************************************

Const PROGRAM_FILES = &H26&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(PROGRAM_FILES)
Set objFolderItem = objFolder.Self

Resultat = objFolderItem.Path & vbcr

Set colItems = objFolder.Items
For Each objItem in colItems
    Resultat = Resultat & objItem.Name & vbCR
Next
MsgBox Resultat,,"Applications installées"
