---
title: ShapeNode.EditingType Property (PowerPoint)
keywords: vbapp10.chm561002
f1_keywords:
- vbapp10.chm561002
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeNode.EditingType
ms.assetid: 6d7f285c-06a2-a9e6-dc3c-bddb1146640f
ms.date: 06/08/2017
---


# ShapeNode.EditingType Property (PowerPoint)

If the specified node is a vertex, this property returns a value that indicates how changes made to the node affect the two segments connected to the node. If the node is a control point for a curved segment, this property returns the editing type of the adjacent vertex. Read-only.


## Syntax

 _expression_. **EditingType**

 _expression_ A variable that represents an **ShapeNode** object.


### Return Value

MsoEditingType


## Remarks

This property is read-only. Use the  **[SetEditingType](shapenodes-seteditingtype-method-powerpoint.md)** method to set the value of this property.

The value of the  **EditingType** property can be one of these **MsoEditingType** constants.


||
|:-----|
|**msoEditingAuto**|
|**msoEditingCorner**|
|**msoEditingSmooth**|
|**msoEditingSymmetric**|

## Example

This example changes all corner nodes to smooth nodes in shape three on  `myDocument`. Shape three must be a freeform drawing.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Nodes

    For n = 1 to .Count

        If .Item(n).EditingType = msoEditingCorner Then

            .SetEditingType n, msoEditingSmooth

        End If

    Next

End With
```


## See also


#### Concepts


[ShapeNode Object](shapenode-object-powerpoint.md)

