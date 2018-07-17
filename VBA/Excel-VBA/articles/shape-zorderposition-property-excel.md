---
title: Shape.ZOrderPosition Property (Excel)
keywords: vbaxl10.chm636116
f1_keywords:
- vbaxl10.chm636116
ms.prod: excel
api_name:
- Excel.Shape.ZOrderPosition
ms.assetid: aaf86516-bf5d-bdb5-1d88-eb1784f9b26f
ms.date: 06/08/2017
---


# Shape.ZOrderPosition Property (Excel)

Returns the position of the specified shape in the z-order. Read-only  **Long** .Read-only


## Syntax

 _expression_ . **ZOrderPosition**

 _expression_ A variable that represents a **Shape** object.


## Remarks

To set the shape's position in the z-order, use the  **[ZOrder](shape-zorder-method-excel.md)** method.

A shape's position in the z-order corresponds to the shape's index number in the  **Shapes** collection. For example, if there are four shapes on `myDocument`, the expression  `myDocument.Shapes(1)` returns the shape at the back of the z-order, and the expression `myDocument.Shapes(4)` returns the shape at the front of the z-order.

Whenever you add a new shape to a collection, it's added to the front of the z-order by default.


## Example

This example adds an oval to  `myDocument` and then places the oval second from the back in the z-order if there is at least one other shape on the document.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.AddShape(msoShapeOval, 100, 100, 100, 300) 
 While .ZOrderPosition > 2 
 .ZOrder msoSendBackward 
 Wend 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

