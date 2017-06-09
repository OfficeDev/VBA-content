---
title: ShapeRange.Adjustments Property (PowerPoint)
keywords: vbapp10.chm548015
f1_keywords:
- vbapp10.chm548015
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Adjustments
ms.assetid: e76f2051-c362-9848-59be-fc3c9662e3a8
ms.date: 06/08/2017
---


# ShapeRange.Adjustments Property (PowerPoint)

Returns an  **[Adjustments](adjustments-object-powerpoint.md)** object that contains adjustment values for all the adjustments in the specified shape. Applies to any **ShapeRange** object that represents an AutoShape, WordArt, or a connector. Read-only.


## Syntax

 _expression_. **Adjustments**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

Adjustments


## Example

This example sets to 0.25 the value of adjustment one for shape three on  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(3).Adjustments(1) = 0.25
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

