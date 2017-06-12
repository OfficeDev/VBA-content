---
title: ShapeNodes Object (Word)
keywords: vbawd10.chm2510
f1_keywords:
- vbawd10.chm2510
ms.prod: word
ms.assetid: f2e13db2-102f-1a14-fd7a-d179f63e513e
ms.date: 06/08/2017
---


# ShapeNodes Object (Word)

A collection of all the  **[ShapeNode](shapenode-object-word.md)** objects in the specified freeform. Each **ShapeNode** object represents either a node between segments in a freeform or a control point for a curved segment of a freeform.


## Remarks

You can create a freeform manually or by using the  **BuildFreeform** and **ConvertToShape** methods.

Use the  **Nodes** property to return the **ShapeNodes** collection. The following example deletes node four in shape three on the active document. For this example to work, shape three must be a freeform with at least four nodes.




```vb
ActiveDocument.Shapes(3).Nodes.Delete 4
```

Use the  **Insert** method to create a new node and add it to the **ShapeNodes** collection. The following example adds a smooth node with a curved segment after node four in shape three on the active document. For this example to work, shape three must be a freeform with at least four nodes.




```vb
With ActiveDocument.Shapes(3).Nodes 
 .Insert 4, msoSegmentCurve, msoEditingSmooth, 210, 100 
End With
```

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

