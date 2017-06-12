---
title: Shape.AlternativeText Property (PowerPoint)
keywords: vbapp10.chm547058
f1_keywords:
- vbapp10.chm547058
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.AlternativeText
ms.assetid: 0ffde7b0-8a91-5456-e092-379491327a15
ms.date: 06/08/2017
---


# Shape.AlternativeText Property (PowerPoint)

Returns or sets the alternative text associated with a shape in a Web presentation. Read/write.


## Syntax

 _expression_. **AlternativeText**

 _expression_ A variable that represents a **Shape** object.


### Return Value

String


## Example

The following example sets the alternative text for the selected shape in the active window. The selected shape is a picture of a mallard duck.


```vb
ActiveWindow.Selection.ShapeRange.AlternativeText = "This is a mallard duck."
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

