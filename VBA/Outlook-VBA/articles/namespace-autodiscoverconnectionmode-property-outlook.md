---
title: NameSpace.AutoDiscoverConnectionMode Property (Outlook)
keywords: vbaol11.chm3303
f1_keywords:
- vbaol11.chm3303
ms.prod: outlook
api_name:
- Outlook.NameSpace.AutoDiscoverConnectionMode
ms.assetid: a73a71ca-0f40-3c7e-bb89-9d6a45775c6f
ms.date: 06/08/2017
---


# NameSpace.AutoDiscoverConnectionMode Property (Outlook)

Returns an  **[OlAutoDiscoverConnectionMode](olautodiscoverconnectionmode-enumeration-outlook.md)** constant that specifies the type of connection for auto-discovery of the Microsoft Exchange server that hosts the primary Exchange account. Read-only.


## Syntax

 _expression_ . **AutoDiscoverConnectionMode**

 _expression_ A variable that represents a **[NameSpace](namespace-object-outlook.md)** object.


## Remarks

This property is similar to the  **[AutoDiscoverConnectionMode](account-autodiscoverconnectionmode-property-outlook.md)** property of the **[Account](account-object-outlook.md)** object. If there are multiple Exchange accounts defined in the current profile, use the **AutoDiscoverConnectionMode** property for the specific account.


## Example

 **NameSpace.AutoDiscoverXml** is an XML string that is returned from the auto-discovery service of the Exchange server. The following code sample uses the **AutoDiscoverConnectionMode** property to show when this XML string is available during a normal Outlook session.


- When the  **[Application.Startup](application-startup-event-outlook.md)** event occurs, if **AutoDiscoverConnectionMode** is not equal to **olAutoDiscoverConnectionUnknown** .
    
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

