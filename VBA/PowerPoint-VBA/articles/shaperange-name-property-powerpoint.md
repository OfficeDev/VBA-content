---
title: ShapeRange.Name Property (PowerPoint)
keywords: vbapp10.chm548029
f1_keywords:
- vbapp10.chm548029
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Name
ms.assetid: b87c7def-f68d-0dde-e971-2201043f6bfc
ms.date: 06/08/2017
---


# ShapeRange.Name Property (PowerPoint)

When a shape is created, Microsoft PowerPoint automatically assigns it a name in the form  _ShapeType Number_, where _ShapeType_ identifies the type of shape or AutoShape, and _Number_ is an integer that's unique within the collection of shapes on the slide. For example, the automatically generated names of the shapes on a slide could be Placeholder 1, Oval 2, and Rectangle 3. To avoid conflict with automatically assigned names, don't use the form _ShapeType Number_ for user-defined names, where _ShapeType_ is a value that is used for automatically generated names, and _Number_ is any positive integer. A shape range must contain exactly one shape. Read/write.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

String


## Remarks

You can use the object's name in conjunction with the  **Item** method to return a reference to the object if the **Item** method for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, then `.Shapes("Rectangle 2")` will return a reference to that shape.


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

