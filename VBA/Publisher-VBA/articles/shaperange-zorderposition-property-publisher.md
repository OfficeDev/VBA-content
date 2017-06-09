---
title: ShapeRange.ZOrderPosition Property (Publisher)
keywords: vbapb10.chm2293848
f1_keywords:
- vbapb10.chm2293848
ms.prod: publisher
api_name:
- Publisher.ShapeRange.ZOrderPosition
ms.assetid: fc35f4dd-ef31-12e0-82a6-be2d0f765527
ms.date: 06/08/2017
---


# ShapeRange.ZOrderPosition Property (Publisher)

Returns a  **Long** indicating the position of the specified shape or shape range in the z-order. Read-only.


## Syntax

 _expression_. **ZOrderPosition**

 _expression_A variable that represents a  **ShapeRange** object.


## Remarks

A shape's position in the z-order corresponds to the shape's index number in the  **Shapes** collection. For example, if there are four shapes on the page, the expression `ActiveDocument.Pages(1).Shapes(1)` returns the shape at the back of the z-order, and the expression `ActiveDocument.Pages(1).Shapes(4)` returns the shape at the front of the z-order.

Whenever you add a new shape to a collection, it is added to the front of the z-order by default.

To set the shape's position in the z-order, use the  **[ZOrder](shape-zorder-method-publisher.md)** method.


## Example

This example adds an oval to the active publication, and then places the oval second from the back in the z-order if there is at least one other shape on the page.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=100, Top:=100, Width:=100, Height:=300) 
 Do While .ZOrderPosition > 2 
 .ZOrder msoSendBackward 
 Loop 
End With 

```


