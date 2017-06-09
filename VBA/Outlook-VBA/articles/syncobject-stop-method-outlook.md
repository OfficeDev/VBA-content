---
title: SyncObject.Stop Method (Outlook)
keywords: vbaol11.chm109
f1_keywords:
- vbaol11.chm109
ms.prod: outlook
api_name:
- Outlook.SyncObject.Stop
ms.assetid: ce74230f-6da7-953e-5a70-157900f4e84d
ms.date: 06/08/2017
---


# SyncObject.Stop Method (Outlook)

Immediately ends synchronizing a user's folders using the specified  **Send/Receive** group.


## Syntax

 _expression_ . **Stop**

 _expression_ A variable that represents a **SyncObject** object.


## Remarks

This method does not undo any synchronization that has already occurred.


## Example

This Microsoft Visual Basic for Applications (VBA) example displays all the  **Send/Receive** groups set up for the user and starts the synchronization based on the user's response. The subroutine following the one below immediately stops the synchronization. The `syc` variable is declared as a public variable so it can be referenced by both the subroutines.


```vb
Public syc As Outlook.SyncObject 
 
Public Sub Sync() 
 Dim nsp As Outlook.NameSpace 
 Dim sycs As Outlook.SyncObjects 
 Dim i As Integer 
 Dim strPrompt As Integer 
 Set nsp = Application.GetNamespace("MAPI") 
 Set sycs = nsp.SyncObjects 
 For i = 1 To sycs.Count 
Set syc = sycs.Item(i) 
strPrompt = MsgBox("Do you wish to synchronize " &; syc.Name &;"?", vbYesNo) 
If strPrompt = vbYes Then 
 syc.Start 
End If 
 Next 
End Sub 
 
Private Sub StopSync() 
 MsgBox "Synchronization stopped by the user." 
 syc.Stop 
End Sub
```


## See also


#### Concepts


[SyncObject Object](syncobject-object-outlook.md)

