---
title: Shape.Nodes Property (Publisher)
keywords: vbapb10.chm2228293
f1_keywords:
- vbapb10.chm2228293
ms.prod: publisher
api_name:
- Publisher.Shape.Nodes
ms.assetid: a1463ff3-5b75-e4b9-df12-985538713c7c
ms.date: 06/08/2017
---


# Shape.Nodes Property (Publisher)

Returns a  **[ShapeNodes](shapenodes-object-publisher.md)** collection that represents the geometric description of the specified shape. Applies to  **Shape** or **ShapeRange** objects that represent freeform drawings.


## Syntax

 _expression_. **Nodes**

 _expression_A variable that represents a  **Shape** object.


## Example

This example adds a smooth node with a curved segment after node four in shape three on page one. Shape three must be a freeform drawing with at least four nodes.


```vb
With ActiveDocument.Pages(1) _ 
 .Shapes(3).Nodes 
 .Insert Index:=4, SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingSmooth, X1:=210, Y1:=100 
End With
```


