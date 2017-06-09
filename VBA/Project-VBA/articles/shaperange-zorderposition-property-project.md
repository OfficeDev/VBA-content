---
title: ShapeRange.ZOrderPosition Property (Project)
ms.prod: project-server
ms.assetid: d9f0d46f-65b1-bb1f-cb75-ce4d7c3b3ab2
ms.date: 06/08/2017
---


# ShapeRange.ZOrderPosition Property (Project)
Gets the position of the shape range in the z-order. Read-only  **Long**.

## Syntax

 _expression_. **ZOrderPosition**

 _expression_ A variable that represents a **ShapeRange** object.


## Remarks

To set the shape position in the z-order, use the [ZOrder](shape-zorder-method-project.md) method.

The position of a shape in the z-order corresponds to the index number of the shape in the  **Shapes** collection. For example, if there are four shapes in the `myReport` report object, the expression `myReport.Shapes(1)` returns the shape at the back of the z-order, and the expression `myReport.Shapes(4)` returns the shape at the front of the z-order.

When you add a shape to a  **Shapes** collection, the shape is added to the front of the z-order by default.


## Property value

 **INT**


## See also


#### Other resources


[ShapeRange Object](shaperange-object-project.md)
