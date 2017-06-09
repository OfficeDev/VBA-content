---
title: TextRange2.BoundWidth Property (PowerPoint)
ms.assetid: 95d4ef10-0f3e-47d8-bfe4-daf8779c74c9
ms.date: 06/08/2017
ms.prod: powerpoint
---


# TextRange2.BoundWidth Property (PowerPoint)

Gets the width, in points, of the text bounding box for the specified text. Read-only.


## Syntax

 _expression_. **BoundWidth**

 _expression_ An expression that returns a **TextRange2** object.


### Return Value

Single


## Remarks

The text bounding box is not the same as the  **TextFrame** object. The **TextFrame** object represents the container in which the text can reside. The text bounding box represents the perimeter immediately surrounding the text.


## Example

This example adds a rounded rectangle to slide one with the same dimensions as the text bounding box.


```vb
With ActivePresentation.Slides(1).Shapes(1) 
 Set txb = .TextFrame.Text 
 Set roundRect = .AddShape(ppShapeRoundRect, _ 
 txb.BoundLeft, txb.BoundTop, txb.BoundWidth, txb.BoundHeight) 
 roundRect.Fill.Transparency = 0.25 
End With 

```


## See also


#### Concepts


[TextRange2 Object (PowerPoint)](textrange2-object-powerpoint.md)


