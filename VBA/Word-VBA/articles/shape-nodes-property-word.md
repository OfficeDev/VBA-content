---
title: Shape.Nodes Property (Word)
keywords: vbawd10.chm161480820
f1_keywords:
- vbawd10.chm161480820
ms.prod: word
api_name:
- Word.Shape.Nodes
ms.assetid: 90904836-e4c4-bbf5-c306-982c9f839ebe
ms.date: 06/08/2017
---


# Shape.Nodes Property (Word)

Returns a  **[ShapeNodes](shapenodes-object-word.md)** collection that represents the geometric description of the specified shape.


## Syntax

 _expression_ . **Nodes**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Example

This example adds a smooth node with a curved segment after node four in shape three in the active document. Shape three must be a freeform drawing with at least four nodes.


```vb
With ActiveDocument.Shapes(3).Nodes 
 .Insert Index:=4, SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingSmooth, X1:=210, Y1:=100 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

