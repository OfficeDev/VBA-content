---
title: ShapeRange.Nodes Property (Publisher)
keywords: vbapb10.chm2293829
f1_keywords:
- vbapb10.chm2293829
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Nodes
ms.assetid: 513be66c-558c-f5f3-ed89-0ef4bc5a0101
ms.date: 06/08/2017
---


# ShapeRange.Nodes Property (Publisher)

Returns a  **[ShapeNodes](shapenodes-object-publisher.md)** collection that represents the geometric description of the specified shape. Applies to  **Shape** or **ShapeRange** objects that represent freeform drawings.


## Syntax

 _expression_. **Nodes**

 _expression_A variable that represents a  **ShapeRange** object.


## Example

This example adds a smooth node with a curved segment after node four in shape three on page one. Shape three must be a freeform drawing with at least four nodes.


```vb
With ActiveDocument.Pages(1) _ 
 .Shapes(3).Nodes 
 .Insert Index:=4, SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingSmooth, X1:=210, Y1:=100 
End With
```


