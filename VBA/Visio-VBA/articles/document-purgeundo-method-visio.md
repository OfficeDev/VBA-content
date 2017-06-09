---
title: Document.PurgeUndo Method (Visio)
keywords: vis_sdr.chm10516450
f1_keywords:
- vis_sdr.chm10516450
ms.prod: visio
api_name:
- Visio.Document.PurgeUndo
ms.assetid: 04556300-8787-5a04-040c-476d864f682e
ms.date: 06/08/2017
---


# Document.PurgeUndo Method (Visio)

Empties the Microsoft Visio queue of undo actions.


## Syntax

 _expression_ . **PurgeUndo**

 _expression_ A variable that represents a **Document** object.


### Return Value

Nothing


## Remarks

After calling the  **PurgeUndo** method, no operation performed before the call can be reversed.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **PurgeUndo** method to purge the undo list.


```vb
 
Public Sub PurgeUndo_Example() 
 
 Application.PurgeUndo 
 
End Sub
```


