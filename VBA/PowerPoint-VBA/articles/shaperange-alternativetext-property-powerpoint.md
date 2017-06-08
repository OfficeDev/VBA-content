---
title: ShapeRange.AlternativeText Property (PowerPoint)
keywords: vbapp10.chm548067
f1_keywords:
- vbapp10.chm548067
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.AlternativeText
ms.assetid: 5255de02-810d-981b-4b2d-9a28fbcdae4c
ms.date: 06/08/2017
---


# ShapeRange.AlternativeText Property (PowerPoint)

Returns or sets the alternative text associated with a shape in a Web presentation. Read/write.


## Syntax

 _expression_. **AlternativeText**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

String


## Example

The following example sets the alternative text for the selected shape in the active window. The selected shape is a picture of a mallard duck.


```vb
ActiveWindow.Selection.ShapeRange.AlternativeText = "This is a mallard duck."
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

