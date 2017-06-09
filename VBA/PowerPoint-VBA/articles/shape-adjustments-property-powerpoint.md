---
title: Shape.Adjustments Property (PowerPoint)
keywords: vbapp10.chm547015
f1_keywords:
- vbapp10.chm547015
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Adjustments
ms.assetid: 2bb29847-cbeb-891b-c1e2-18e8ea7e8122
ms.date: 06/08/2017
---


# Shape.Adjustments Property (PowerPoint)

Returns an  **[Adjustments](adjustments-object-powerpoint.md)** object that contains adjustment values for all the adjustments in the specified shape. Applies to any **Shape** object that represents an AutoShape, WordArt, or a connector. Read-only.


## Syntax

 _expression_. **Adjustments**

 _expression_ A variable that represents a **Shape** object.


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


[Shape Object](shape-object-powerpoint.md)

