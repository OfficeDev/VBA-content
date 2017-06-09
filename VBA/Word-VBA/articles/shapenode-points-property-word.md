---
title: ShapeNode.Points Property (Word)
keywords: vbawd10.chm164429925
f1_keywords:
- vbawd10.chm164429925
ms.prod: word
api_name:
- Word.ShapeNode.Points
ms.assetid: 2d64956f-1ba5-66d0-c4db-cf54c594ca0c
ms.date: 06/08/2017
---


# ShapeNode.Points Property (Word)

Returns the position of the specified node as a coordinate pair. Read-only  **Variant** .


## Syntax

 _expression_ . **Points**

 _expression_ A variable that represents a **[ShapeNode](shapenode-object-word.md)** object.


## Remarks

Each coordinate is expressed in points. Use the  **[SetPosition](shapenodes-setposition-method-word.md)** method to set the location of the node.


## Example

This example moves node two in shape three on myDocument to the right 200 points and down 300 points. Shape three must be a freeform drawing.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes(3).Nodes 
 pointsArray = .Item(2).Points 
 currXvalue = pointsArray(1, 1) 
 currYvalue = pointsArray(1, 2) 
 .SetPosition 2, currXvalue + 200, currYvalue + 300 
End With
```


## See also


#### Concepts


[ShapeNode Object](shapenode-object-word.md)

