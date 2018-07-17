---
title: Outlook COM add-in template
keywords: vbaol11.chm5268747
f1_keywords:
- vbaol11.chm5268747
ms.prod: outlook
ms.assetid: 6c6b4f10-2d7d-75bc-8a0c-6888b560e569
ms.date: 06/08/2017
---


# Outlook COM add-in template

The following code example provides the empty event procedures required to implement a COM add-in.


```vb
Implements IDTExtensibility2 
 
Private Sub IDTExtensibility2_OnAddInsUpdate(custom() As Variant) 
' Occurs when the set of connected COM add-ins changes, that is when 
' any other add-in is connected or disconnected. 
' The custom argument is ignored. 
End Sub 
 
Private Sub IDTExtensibility2_OnBeginShutdown(custom() As Variant) 
' If the COM add-in is connected, occurs when Outlook begins its 
' shutdown routines. 
' The custom argument is ignored. 
 
End Sub 
 
Private Sub IDTExtensibility2_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant) 
' Occurs when the COM add-in is connected. 
' The Application argument is the Outlook Application object. 
' The ConnectMode argument specifies how the COM add-in was connected. 
' It can be 
' ext_cm_AfterStartup Add-in was connected after Outlook started, 
' or the Connect property of the corresponding 
' COMAddIn object was set to True 
' ext_cm_Startup Add-in was connected on startup 
' ext_cm_External 
' ext_cm_CommandLine 
' The AddInInst argument is the COMAddIn object that refers to the current 
' instance of the add-in itself. 
' The custom argument is ignored. 
End Sub 
 
Private Sub IDTExtensibility2_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant) 
' Occurs when the COM add-in is disconnected. 
' The RemoveMode argument specifies how the COM add-in was disconnected. 
' It can be 
' ext_dm_HostShutdown Add-in was disconnected when Outlook was 
' closed. 
' ext_dm_UserClosed Add-in was disconnected when the user 
' cleared the corresponding check box in the 
' COM Add-ins dialog box, or the Connect 
' property of the corresponding COMAddIn 
' object was set to False. 
' The custom argument is ignored. 
End Sub 
 
Private Sub IDTExtensibility2_OnStartupComplete(custom() As Variant) 
' If the COM add-in connects at startup, occurs when Outlook completes 
' its startup routines. This event does not occur if the COM add-in is not 
' connected when Outlook loads, even when the user connects the add-in in 
' the COM Add-ins dialog box. 
' The custom argument is ignored. 
End Sub
```


