---
title: Application.NewMail Event (Outlook)
keywords: vbaol11.chm430
f1_keywords:
- vbaol11.chm430
ms.prod: outlook
api_name:
- Outlook.Application.NewMail
ms.assetid: cfc848e8-98b1-163a-c177-53993c20bb14
ms.date: 06/08/2017
---


# Application.NewMail Event (Outlook)

Occurs when one or more new e-mail messages are received in the  **Inbox**. 


## Syntax

 _expression_ . **NewMail**

 _expression_ A variable that represents an **[Application](application-object-outlook.md)** object.


## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).

The  **NewMail** event fires when new messages arrive in the Inbox and before client rule processing occurs. If you want to process items that arrive in the **Inbox**, consider using the  **[ItemAdd](items-itemadd-event-outlook.md)** event on the collection of items in the **Inbox**. The  **ItemAdd** event passes a reference to each item that is added to a folder.

The  **NewMail** event does not fire when the user is in offline mode.


## Example

This Microsoft Visual Basic for Applications (VBA) example displays the  **Inbox** folder when a new e-mail message arrives. The sample code must be placed in a class module, and the `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents myOlApp As Outlook.Application 
 
 
 
Sub Initialize_handler() 
 
 Set myOlApp = Outlook.Application 
 
End Sub 
 
 
 
Private Sub myOlApp_NewMail() 
 
 Dim myExplorers As Outlook.Explorers 
 
 Dim myFolder As Outlook.Folder 
 
 Dim x As Integer 
 
 Set myExplorers = myOlApp.Explorers 
 
 Set myFolder = myOlApp.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox) 
 
 If myExplorers.Count <> 0 Then 
 
 For x = 1 To myExplorers.Count 
 
 On Error GoTo skipif 
 
 If myExplorers.Item(x).CurrentFolder.Name = "Inbox" Then 
 
 myExplorers.Item(x).Display 
 
 myExplorers.Item(x).Activate 
 
 Exit Sub 
 
 End If 
 
skipif: 
 
 Next x 
 
 End If 
 
 On Error GoTo 0 
 
 myFolder.Display 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-outlook.md)

