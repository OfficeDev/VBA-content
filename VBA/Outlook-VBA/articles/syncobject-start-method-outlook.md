---
title: SyncObject.Start Method (Outlook)
keywords: vbaol11.chm108
f1_keywords:
- vbaol11.chm108
ms.prod: outlook
api_name:
- Outlook.SyncObject.Start
ms.assetid: 3e826228-b8a4-42df-1757-3248acd26a2b
ms.date: 06/08/2017
---


# SyncObject.Start Method (Outlook)

Begins synchronizing a user's folders using the specified  **Send\Receive** group.


## Syntax

 _expression_ . **Start**

 _expression_ An expression that returns a **[SyncObject](syncobject-object-outlook.md)** object.


## Example

This Microsoft Visual Basic for Applications (VBA) example displays all the  **Send\Receive** groups set up for the user and starts the synchronization based on user's response.


```vb
Public Sub Sync() 
 Dim nsp As Outlook.NameSpace 
 Dim sycs As Outlook.SyncObjects 
 Dim syc As Outlook.SyncObject 
 Dim i As Integer 
 Dim strPrompt As Integer 
 Set nsp = Application.GetNamespace("MAPI") 
 Set sycs = nsp.SyncObjects 
 For i = 1 To sycs.Count 
Set syc = sycs.Item(i) 
strPrompt = MsgBox( _ 
 "Do you wish to synchronize " &; syc.Name &;"?", vbYesNo) 
If strPrompt = vbYes Then 
 syc.Start 
End If 
 Next 
End Sub
```


## See also


#### Concepts


[SyncObject Object](syncobject-object-outlook.md)

