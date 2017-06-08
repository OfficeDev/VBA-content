---
title: Shape.Nodes Property (PowerPoint)
keywords: vbapp10.chm547030
f1_keywords:
- vbapp10.chm547030
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Nodes
ms.assetid: 85021d71-78f8-43e5-5a15-a0c1ae29ef61
ms.date: 06/08/2017
---


# Shape.Nodes Property (PowerPoint)

Returns a  **[ShapeNodes](shapenodes-object-powerpoint.md)** collection that represents the geometric description of the specified shape. Applies to **Shape** objects that represent freeform drawings.


## Syntax

 _expression_. **Nodes**

 _expression_ A variable that represents a **Shape** object.


## Example

This example adds a smooth node with a curved segment after node four in shape three on  `myDocument`. Shape three must be a freeform drawing with at least four nodes.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Nodes
    .Insert Index:=4, SegmentType:=msoSegmentCurve, _
        EditingType:=msoEditingSmooth, X1:=210, Y1:=100
End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

