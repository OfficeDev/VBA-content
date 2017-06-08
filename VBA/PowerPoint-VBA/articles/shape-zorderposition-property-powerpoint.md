---
title: Shape.ZOrderPosition Property (PowerPoint)
keywords: vbapp10.chm547043
f1_keywords:
- vbapp10.chm547043
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.ZOrderPosition
ms.assetid: 07e6c756-ae7b-f6d9-d903-92aef3b7fa9e
ms.date: 06/08/2017
---


# Shape.ZOrderPosition Property (PowerPoint)

Returns the position of the specified shape in the z-order. Read-only.


## Syntax

 _expression_. **ZOrderPosition**

 _expression_ A variable that represents a **Shape** object.


### Return Value

Long


## Remarks

 `Shapes(1)` returns the shape at the back of the z-order, and `Shapes(Shapes.Count)` returns the shape at the front of the z-order.

To set the shape's position in the z-order, use the  **ZOrder** method.

A shape's position in the z-order corresponds to the shape's index number in the  **Shapes** collection. For example, if there are four shapes on the slide, the expression `myDocument.Shapes(1)` returns the shape at the back of the z-order, and the expression `myDocument.Shapes(4)` returns the shape at the front of the z-order.

Whenever you add a new shape to a collection, it is added to the front of the z-order by default.


## Example

This example adds an oval to  `myDocument` and then places the oval second from the back in the z-order if there is at least one other shape on the slide.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300)

    While .ZOrderPosition > 2

        .ZOrder msoSendBackward

    Wend

End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

