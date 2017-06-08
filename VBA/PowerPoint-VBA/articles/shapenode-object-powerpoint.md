---
title: ShapeNode Object (PowerPoint)
keywords: vbapp10.chm561000
f1_keywords:
- vbapp10.chm561000
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeNode
ms.assetid: 031edfef-4eae-39b2-0c73-90d2065741aa
ms.date: 06/08/2017
---


# ShapeNode Object (PowerPoint)

Represents the geometry and the geometry-editing properties of the nodes in a user-defined freeform.


## Remarks

 Nodes include the vertices between the segments of the freeform and the control points for curved segments. The **ShapeNode** object is a member of the **[ShapeNodes](shapenodes-object-powerpoint.md)** collection. The **ShapeNodes** collection contains all the nodes in a freeform.


## Example

Use  **Nodes** (index), where index is the node index number, to return a single **ShapeNode** object. If node one in shape three on `myDocument` is a corner point, the following example makes it a smooth point. For this example to work, shape three must be a freeform.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3)

    If .Nodes(1).EditingType = msoEditingCorner Then

        .Nodes.SetEditingType 1, msoEditingSmooth

    End If

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

