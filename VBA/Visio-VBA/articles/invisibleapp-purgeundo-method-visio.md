---
title: InvisibleApp.PurgeUndo Method (Visio)
keywords: vis_sdr.chm17516450
f1_keywords:
- vis_sdr.chm17516450
ms.prod: visio
api_name:
- Visio.InvisibleApp.PurgeUndo
ms.assetid: 8f1ed9a6-1e1e-0059-d0df-1b628e0f45ff
ms.date: 06/08/2017
---


# InvisibleApp.PurgeUndo Method (Visio)

Empties the Microsoft Visio queue of undo actions.


## Syntax

 _expression_ . **PurgeUndo**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Nothing


## Remarks

After calling the  **PurgeUndo** method, no operation performed before the call can be reversed.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **PurgeUndo** method to clear the undo list.


```vb
 
Public Sub PurgeUndo_Example() 
 
 Application.PurgeUndo 
 
End Sub
```


