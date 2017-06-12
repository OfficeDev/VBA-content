---
title: Retrieve a List of Installed Printers
ms.prod: access
ms.assetid: e3162c3e-6b5b-77c3-32f9-1fdfa64cdefc
ms.date: 06/08/2017
---


# Retrieve a List of Installed Printers

You use the  **[Printers](application-printers-property-access.md)** property of the **[Application](application-object-access.md)** object to return the **[Printers](printers-object-access.md)** collection. The following procedure illustrates how to enumerate through each **Printer** object in the **Printers** collection by using a **For Each…Next** statement. A message box is displayed with information about each installed printer.


```vb
Sub ShowPrinters() 
    Dim strCount As String 
    Dim strMsg As String 
    Dim prtLoop As Printer 
     
    On Error GoTo ShowPrinters_Err 
 
    If Printers.Count > 0 Then 
        ' Get count of installed printers. 
        strMsg = "Printers installed: " &; Printers.Count &; vbCrLf &; vbCrLf 
     
        ' Enumerate printer system properties. 
        For Each prtLoop In Application.Printers 
            With prtLoop 
                strMsg = strMsg _ 
                    &; "Device name: " &; .DeviceName &; vbCrLf _ 
                    &; "Driver name: " &; .DriverName &; vbCrLf _ 
                    &; "Port: " &; .Port &; vbCrLf &; vbCrLf 
            End With 
        Next prtLoop 
     
    Else 
        strMsg = "No printers are installed." 
    End If 
     
    ' Display printer information. 
    MsgBox Prompt:=strMsg, Buttons:=vbOKOnly, Title:="Installed Printers" 
     
ShowPrinters_End: 
    Exit Sub 
     
ShowPrinters_Err: 
    MsgBox Prompt:=Err.Description, Buttons:=vbCritical &; vbOKOnly, _ 
        Title:="Error Number " &; Err.Number &; " Occurred" 
    Resume ShowPrinters_End 
     
End Sub
```


