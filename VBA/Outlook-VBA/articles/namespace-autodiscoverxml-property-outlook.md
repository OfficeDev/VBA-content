---
title: NameSpace.AutoDiscoverXml Property (Outlook)
keywords: vbaol11.chm3263
f1_keywords:
- vbaol11.chm3263
ms.prod: outlook
api_name:
- Outlook.NameSpace.AutoDiscoverXml
ms.assetid: 34834000-1f53-2bfb-7546-886c6e2716fd
ms.date: 06/08/2017
---


# NameSpace.AutoDiscoverXml Property (Outlook)

Returns a  **String** that represents information in XML retrieved from the auto-discovery service for the Microsoft Exchange server that hosts the primary Exchange account. Read-only.


## Syntax

 _expression_ . **AutoDiscoverXml**

 _expression_ A variable that represents a **[NameSpace](namespace-object-outlook.md)** object.


## Remarks

This property is similar to the  **[AutoDiscoverXml](account-autodiscoverxml-property-outlook.md)** property of the **[Account](account-object-outlook.md)** object. If there are multiple Exchange accounts defined in the current profile, use the **AutoDiscoverXml** property for the specific account.

The returned string of XML contains information about various Web services (for example, availability service and unified messaging service) and available servers. 

An error is returned if the active profile does not contain an account that is connected to a server that is Microsoft Exchange Server 2007 or later.


## Example

 **NameSpace.AutoDiscoverXml** is an XML string that is returned from the auto-discovery service of the Exchange server. The following code sample uses the **AutoDiscoverConnectionMode** property to show when this XML string is available during a normal Outlook session.


- When the  **[Application.Startup](application-startup-event-outlook.md)** event occurs, if **[NameSpace.AutoDiscoverConnectionMode](namespace-autodiscoverconnectionmode-property-outlook.md)** is not equal to **olAutoDiscoverConnectionUnknown** .
    
- When the  **[NameSpace.AutoDiscoverComplete](namespace-autodiscovercomplete-event-outlook.md)** event occurs, if **AutoDiscoverConnectionMode** is not equal to **olAutoDiscoverConnectionUnknown** .
    





```vb
Dim WithEvents Session As NameSpace 
 
Dim LastAutoDiscoverXml As String 
 
Dim LastAutoDiscoverConnectionMode As OlAutoDiscoverConnectionMode 
 
 
 
Private Sub Application_Startup() 
 
 Set Session = Application.Session 
 
 If (Session.AutoDiscoverConnectionMode <> olAutoDiscoverConnectionUnknown) Then 
 
 LastAutoDiscoverXml = Session.AutoDiscoverXml 
 
 LastAutoDiscoverConnectionMode = Session.AutoDiscoverConnectionMode 
 
 DoAutoDiscoverBasedWork 
 
 End If 
 
End Sub 
 
 
 
Private Sub Session_AutoDiscoverComplete() 
 
 LastAutoDiscoverXml = Session.AutoDiscoverXml 
 
 LastAutoDiscoverConnectionMode = Session.AutoDiscoverConnectionMode 
 
 If LastAutoDiscoverConnectionMode <> olAutoDiscoverConnectionUnknown Then 
 
 DoAutoDiscoverBasedWork 
 
 End If 
 
End Sub 
 
 
 
Private Sub DoAutoDiscoverBasedWork() 
 
 ' Do activity requires auto discover information 
 
 Dim displayName As String 
 
 Dim posStartTag, posEndTag As Integer 
 
 posStartTag = InStr(1, LastAutoDiscoverXml, "<DisplayName>") 
 
 posEndTag = InStr(1, LastAutoDiscoverXml, "</DisplayName>") 
 
 
 
 If (posStartTag > 1 And posEndTag > 1) Then 
 
 displayName = Mid(LastAutoDiscoverXml, posStartTag + 13, posEndTag - posStartTag - 13) 
 
 Debug.Print "DisplayName = " &; displayName 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

