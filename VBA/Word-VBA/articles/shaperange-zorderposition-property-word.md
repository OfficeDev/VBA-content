---
title: ShapeRange.ZOrderPosition Property (Word)
keywords: vbawd10.chm162857089
f1_keywords:
- vbawd10.chm162857089
ms.prod: word
api_name:
- Word.ShapeRange.ZOrderPosition
ms.assetid: 358898b6-eab7-96f8-8f4e-bb993a425e5e
ms.date: 06/08/2017
---


# ShapeRange.ZOrderPosition Property (Word)

Returns a  **Long** that represents the position of the specified shape in the z-order. Read-only.


## Syntax

 _expression_ . **ZOrderPosition**

 _expression_ An expression that returns a **ShapeRange** object.


## Remarks

 `Shapes(1)` returns the shape at the back of the z-order, and `Shapes(Shapes.Count)` returns the shape at the front of the z-order. This property is read-only. To set the shape's position in the z-order, use the **ZOrder** method.

A shape's position in the z-order corresponds to the shape's index number in the Shapes collection. For example, if there are four shapes on myDocument, the expression  `myDocument.Shapes(1)` returns the shape at the back of the z-order, and the expression `myDocument.Shapes(4)` returns the shape at the front of the z-order.

Whenever you add a new shape to a collection, it is added to the front of the z-order by default.


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

