---
title: Shape.Name Property (PowerPoint)
keywords: vbapp10.chm547029
f1_keywords:
- vbapp10.chm547029
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Name
ms.assetid: 66e1d7e8-9398-8f01-d130-7206a48a63b3
ms.date: 06/08/2017
---


# Shape.Name Property (PowerPoint)

When a shape is created, Microsoft PowerPoint automatically assigns it a name in the form  _ShapeType Number_, where _ShapeType_ identifies the type of shape or AutoShape, and _Number_ is an integer that's unique within the collection of shapes on the slide. For example, the automatically generated names of the shapes on a slide could be Placeholder 1, Oval 2, and Rectangle 3. To avoid conflict with automatically assigned names, don't use the form _ShapeType Number_ for user-defined names, where _ShapeType_ is a value that is used for automatically generated names, and _Number_ is any positive integer. A shape range must contain exactly one shape. Read/write.


## Syntax

 _expression_. **Name**

 _expression_ A variable that represents a **Shape** object.


### Return Value

String


## Remarks

You can use the object's name in conjunction with the  **Item** method to return a reference to the object if the **Item** method for the collection that contains the object takes a **Variant** argument. For example, if the value of the **Name** property for a shape is Rectangle 2, then `.Shapes("Rectangle 2")` will return a reference to that shape.


## Example

This example sets the name of object two on slide one in the active presentation to "big triangle."


```vb
ActivePresentation.Slides(1).Shapes(2).Name = "big triangle"
```

This example sets the fill color for the shape named "big triangle" on slide one in the active presentation.




```vb
ActivePresentation.Slides(1) _
    .Shapes("big triangle").Fill.ForeColor.RGB = RGB(0, 0, 255)
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

