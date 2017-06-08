---
title: Shape.ZOrderPosition Property (Word)
keywords: vbawd10.chm161480833
f1_keywords:
- vbawd10.chm161480833
ms.prod: word
api_name:
- Word.Shape.ZOrderPosition
ms.assetid: a1335280-721a-7746-b8e5-b61cf5b8a1e2
ms.date: 06/08/2017
---


# Shape.ZOrderPosition Property (Word)

Returns a  **Long** that represents the position of the specified shape in the z-order. Read-only.


## Syntax

 _expression_ . **ZOrderPosition**

 _expression_ An expression that returns a **Shape** object.


## Remarks

 `Shapes(1)` returns the shape at the back of the z-order, and `Shapes(Shapes.Count)` returns the shape at the front of the z-order. This property is read-only. To set the shape's position in the z-order, use the **ZOrder** method.

A shape's position in the z-order corresponds to the shape's index number in the Shapes collection. For example, if there are four shapes on myDocument, the expression  `myDocument.Shapes(1)` returns the shape at the back of the z-order, and the expression `myDocument.Shapes(4)` returns the shape at the front of the z-order.

Whenever you add a new shape to a collection, it is added to the front of the z-order by default.


## Example

This example adds an oval to myDocument and then places the oval second from the back in the z-order if there is at least one other shape on the document.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300) 
 While .ZOrderPosition > 2 
 .ZOrder msoSendBackward 
 Wend 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

