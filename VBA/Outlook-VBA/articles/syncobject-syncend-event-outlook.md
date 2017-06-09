---
title: SyncObject.SyncEnd Event (Outlook)
keywords: vbaol11.chm114
f1_keywords:
- vbaol11.chm114
ms.prod: outlook
api_name:
- Outlook.SyncObject.SyncEnd
ms.assetid: 6e36b438-bbd3-4810-f072-7b669c308bc6
ms.date: 06/08/2017
---


# SyncObject.SyncEnd Event (Outlook)

Occurs immediately after Microsoft Outlook finishes synchronizing a user's folders using the specified  **Send/Receive** group.


## Syntax

 _expression_ . **SyncEnd**

 _expression_ A variable that represents a **SyncObject** object.


## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Microsoft Visual Basic for Applications (VBA) example displays a message when synchronization is complete. The sample code must be placed in a class module, and the  `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook.


```vb
Dim WithEvents mySync As Outlook.SyncObject 
 
Sub Initialize_handler() 
 Set mySync = Application.Session.SyncObjects.Item(1) 
 mySync.Start 
End Sub 
 
Private Sub mySync_SyncEnd() 
 MsgBox "Synchronization is complete." 
End Sub
```


## See also


#### Concepts


[SyncObject Object](syncobject-object-outlook.md)

