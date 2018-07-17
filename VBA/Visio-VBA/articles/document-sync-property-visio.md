---
title: Document.Sync Property (Visio)
keywords: vis_sdr.chm10560137
f1_keywords:
- vis_sdr.chm10560137
ms.prod: visio
api_name:
- Visio.Document.Sync
ms.assetid: 1e5ef6da-a665-024f-5e35-e8518f4d1054
ms.date: 06/08/2017
---


# Document.Sync Property (Visio)

Returns a Microsoft Office  **Sync** object that provides information about the status of the active document in a shared workspace and the ability to perform a set of actions. Read-only.


## Syntax

 _expression_ . **Sync**

 _expression_ A variable that represents a **Document** object.


### Return Value

Object


## Remarks

If the  **Sync** object is unavailable because the synchronization engine fails to respond, the following error message is displayed: "The synchronization engine is not available."


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Sync** property to get a **Sync** object and get the status of the active document in a shared workspace.


```vb
Public Sub Sync_Example 
 
 Dim vsoSync As Sync 
 Dim currentStatus As Integer 
 
 Set vsoSync = ActiveDocument.Sync 
 currentStatus = vsoSync.Status 
 
 If currentStatus = msoSyncStatusLatest 
 
 Msgbox "You have the most up-to-date copy." 
 
 Else 
 
 Msgbox "You need to update." 
 
 End if 
 
End Sub
```


