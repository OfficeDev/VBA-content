---
title: Application.PurgeUndo Method (Visio)
keywords: vis_sdr.chm10016450
f1_keywords:
- vis_sdr.chm10016450
ms.prod: visio
api_name:
- Visio.Application.PurgeUndo
ms.assetid: d5d18607-2b1d-6b47-2a81-43345ff0be8a
ms.date: 06/08/2017
---


# Application.PurgeUndo Method (Visio)

Empties the Microsoft Visio queue of undo actions.


## Syntax

 _expression_ . **PurgeUndo**

 _expression_ A variable that represents an **Application** object.


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


