---
title: DropCap.Clear Method (Publisher)
keywords: vbapb10.chm5505042
f1_keywords:
- vbapb10.chm5505042
ms.prod: publisher
api_name:
- Publisher.DropCap.Clear
ms.assetid: 7c30e774-c520-076a-41d8-7c68679f58bc
ms.date: 06/08/2017
---


# DropCap.Clear Method (Publisher)

Removes the dropped capital letter formatting.


## Syntax

 _expression_. **Clear**

 _expression_A variable that represents a  **DropCap** object.


## Example

This example removes the dropped capital letter formatting in the specified text frame.


```vb
Sub ClearDropCap() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame _ 
 .TextRange.DropCap.Clear 
End Sub
```


