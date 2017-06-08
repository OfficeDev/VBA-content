---
title: ShapeNode Object (Word)
keywords: vbawd10.chm2509
f1_keywords:
- vbawd10.chm2509
ms.prod: word
api_name:
- Word.ShapeNode
ms.assetid: d5afb71a-a218-57f3-87f0-171094ba6610
ms.date: 06/08/2017
---


# ShapeNode Object (Word)

Represents the geometry and the geometry-editing properties of the nodes in a user-defined freeform. Nodes include the vertices between the segments of the freeform and the control points for curved segments. The  **ShapeNode** object is a member of the **ShapeNodes** collection. The **[ShapeNodes](shapenodes-object-word.md)** collection contains all the nodes in a freeform.


## Remarks

Use  **Nodes** (Index), where Index is the node index number, to return a single **ShapeNode** object. If node one in shape three on the active document is a corner point, the following example makes it a smooth point. For this example to work, shape three must be a freeform.


```vb
With ActiveDocument.Shapes(3) 
 If .Nodes(1).EditingType = msoEditingCorner Then 
 .Nodes.SetEditingType 1, msoEditingSmooth 
 End If 
End With
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


