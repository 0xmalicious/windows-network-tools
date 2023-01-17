Option Explicit
Dim strComputer,colAdapters,objAdapter,objWMIService
Dim objItem,colItems,MyIP,ComputerName,i
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" _
& strComputer & "\root\cimv2")
 
Set colAdapters = objWMIService.ExecQuery _
("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
 
Set colItems = objWMIService.ExecQuery _
("Select * from Win32_OperatingSystem")
 
For Each objAdapter in colAdapters
    MyIp = MyIp & objAdapter.IPAddress(i) & VbcrLf & VbTab
Next
 
For Each objItem in colItems
    ComputerName = ComputerName & objItem.CSName
Next
 
Msgbox "Nom de l'ordinateur : " & ComputerName & VbcrLf & VbcrLf &_
"Adresse IP : " & MyIp ,64,"Nom du Computer + Adresses IP(s)"
