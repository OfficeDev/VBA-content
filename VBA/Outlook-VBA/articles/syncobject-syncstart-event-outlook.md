---
title: SyncObject.SyncStart Event (Outlook)
keywords: vbaol11.chm111
f1_keywords:
- vbaol11.chm111
ms.prod: outlook
api_name:
- Outlook.SyncObject.SyncStart
ms.assetid: 225367bc-3bff-cea0-3e8c-71a30256f45d
ms.date: 06/08/2017
---


# SyncObject.SyncStart Event (Outlook)

Occurs when Microsoft Outlook begins synchronizing a user's folders using the specified  **Send\Receive** group.


## Syntax

 _expression_ . **SyncStart**

 _expression_ A variable that represents a **SyncObject** object.


## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Visual Basic for Applications (VBA) example displays a message telling the user that the synchronization might take a long time. The sample code must be placed in a class module, and the  `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Dim WithEvents mySync As Outlook.SyncObject 
 
Sub Initialize_handler() 
 Set mySync = Application.Session.SyncObjects.Item(1) 
 mySync.Start 
End Sub 
 
Private Sub mySync_SyncStart() 
 MsgBox "Synchronization is about to start. It might take a long time to complete." 
End Sub
```


## See also


#### Concepts


[SyncObject Object](syncobject-object-outlook.md)

