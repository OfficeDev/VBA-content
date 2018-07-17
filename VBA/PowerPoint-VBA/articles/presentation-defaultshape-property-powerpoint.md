---
title: Presentation.DefaultShape Property (PowerPoint)
keywords: vbapp10.chm583019
f1_keywords:
- vbapp10.chm583019
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.DefaultShape
ms.assetid: 318ec04a-8b30-29b3-c8a6-732564efd7a8
ms.date: 06/08/2017
---


# Presentation.DefaultShape Property (PowerPoint)

Returns a  **[Shape](shape-object-powerpoint.md)** object that represents the default shape for the presentation. Read-only.


## Syntax

 _expression_. **DefaultShape**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

Shape


## Example

This example adds a shape to slide one in the active presentation, sets the default fill color to red for shapes in the presentation, and then adds another shape. This second shape will automatically have the new default fill color applied to it.


```vb
With Application.ActivePresentation

    Set sld1Shapes = .Slides(1).Shapes

    sld1Shapes.AddShape msoShape16pointStar, 20, 20, 100, 100

    .DefaultShape.Fill.ForeColor.RGB = RGB(255, 0, 0)

    sld1Shapes.AddShape msoShape16pointStar, 150, 20, 100, 100

End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

