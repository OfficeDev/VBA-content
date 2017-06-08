---
title: NameSpace.AutoDiscoverComplete Event (Outlook)
keywords: vbaol11.chm3301
f1_keywords:
- vbaol11.chm3301
ms.prod: outlook
api_name:
- Outlook.NameSpace.AutoDiscoverComplete
ms.assetid: b7cac212-4d38-660e-0caf-48f97035f14a
ms.date: 06/08/2017
---


# NameSpace.AutoDiscoverComplete Event (Outlook)

Occurs after Microsoft Outlook has finished accessing the auto-discovery service of the Microsoft Exchange server that hosts the primary Exchange account and has the related information available in  **[NameSpace.AutoDiscoverXml](namespace-autodiscoverxml-property-outlook.md)** .


## Syntax

 _expression_ . **AutoDiscoverComplete**

 _expression_ A variable that represents a **[NameSpace](namespace-object-outlook.md)** object.


## Remarks

This event is similar to the  **[AutoDiscoverComplete](accounts-autodiscovercomplete-event-outlook.md)** event of the **[Accounts](accounts-object-outlook.md)** object. If there are multiple Exchange accounts defined in the current profile, use the **AutoDiscoverComplete** event of the **Accounts** object that specifies the particular account.


## Example

 **NameSpace.AutoDiscoverXml** is an XML string that is returned from the auto-discovery service of the Exchange server. The following code sample shows when this XML string is available during a normal Outlook session:


1. When the  **[Application.Startup](application-startup-event-outlook.md)** event occurs, if **[NameSpace.AutoDiscoverConnectionMode](namespace-autodiscoverconnectionmode-property-outlook.md)** is not equal to **olAutoDiscoverConnectionUnknown**
    
2. When the  **AutoDiscoverComplete** event occurs, if **AutoDiscoverConnectionMode** is not equal to **olAutoDiscoverConnectionUnknown**
    





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

