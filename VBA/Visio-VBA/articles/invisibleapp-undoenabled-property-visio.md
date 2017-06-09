---
title: InvisibleApp.UndoEnabled Property (Visio)
keywords: vis_sdr.chm17514610
f1_keywords:
- vis_sdr.chm17514610
ms.prod: visio
api_name:
- Visio.InvisibleApp.UndoEnabled
ms.assetid: c3dc1bf4-c3bd-53dd-62e6-f2b6e3f07cb2
ms.date: 06/08/2017
---


# InvisibleApp.UndoEnabled Property (Visio)

Determines whether undo information is maintained in memory. Read/write.


## Syntax

 _expression_ . **UndoEnabled**

 _expression_ A variable that represents an **InvisibleApp** object.


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


