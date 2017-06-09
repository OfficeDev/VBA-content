---
title: Document.UndoEnabled Property (Visio)
keywords: vis_sdr.chm10514610
f1_keywords:
- vis_sdr.chm10514610
ms.prod: visio
api_name:
- Visio.Document.UndoEnabled
ms.assetid: c7164cb6-7ce4-b65d-7f5b-1f3987a3fe21
ms.date: 06/08/2017
---


# Document.UndoEnabled Property (Visio)

Determines whether undo information is maintained in memory. Read/write.


## Syntax

 _expression_ . **UndoEnabled**

 _expression_ A variable that represents a **Document** object.


### Return Value

Boolean


## Remarks

When Microsoft Visio starts, the value of the  **UndoEnabled** property is **True** . Setting the value of the **UndoEnabled** property to **False** discontinues the collection of undo information in memory and clears the existing undo information.

You should attempt to maintain the property at its current value across the complete operation that you perform. In other words, use code structured like this:




```vb
blsPrevious = Application.UndoEnabled 
Application.UndoEnabled = False 
 
'Large operation here 
Application.UndoEnabled = blsPrevious 

```


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **UndoEnabled** method to disable and then re-enable undo behavior in Visio.


```vb
Public Sub UndoEnabled_Example() 
 
 'Disable undo 
 Application.UndoEnabled = False 
 
 'Draw three shapes. 
 ActivePage.DrawRectangle 1, 2, 2, 1 
 ActivePage.DrawOval 3, 4, 4, 3 
 ActivePage.DrawLine 4, 5, 5, 4 
 
 'Enable undo. 
 Application.UndoEnabled = True 
 
End Sub
```


