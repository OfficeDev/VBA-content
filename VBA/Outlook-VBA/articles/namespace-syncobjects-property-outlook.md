---
title: NameSpace.SyncObjects Property (Outlook)
keywords: vbaol11.chm770
f1_keywords:
- vbaol11.chm770
ms.prod: outlook
api_name:
- Outlook.NameSpace.SyncObjects
ms.assetid: 0948f154-022f-b12e-87e3-1b3a4ce127c3
ms.date: 06/08/2017
---


# NameSpace.SyncObjects Property (Outlook)

Returns a  **[SyncObjects](syncobjects-object-outlook.md)** collection containing all Send\Receive groups. Read-only.


## Syntax

 _expression_ . **SyncObjects**

 _expression_ A variable that represents a **NameSpace** object.


## Example

This Microsoft Visual Basic for Applications (VBA) example displays all the Send and Receive groups set up for the user and starts the synchronization based on the user's response.


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
 
 strPrompt = MsgBox("Do you wish to synchronize " &; syc.Name &;"?", vbYesNo) 
 
 If strPrompt = vbYes Then 
 
 syc.Start 
 
 End If 
 
 Next 
 
End Sub
```


## See also


#### Concepts


[NameSpace Object](namespace-object-outlook.md)

