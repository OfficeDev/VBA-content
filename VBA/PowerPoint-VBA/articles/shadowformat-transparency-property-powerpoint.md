---
title: ShadowFormat.Transparency Property (PowerPoint)
keywords: vbapp10.chm554008
f1_keywords:
- vbapp10.chm554008
ms.prod: powerpoint
api_name:
- PowerPoint.ShadowFormat.Transparency
ms.assetid: 5ab7440c-111e-c3a5-159f-abe752893fd3
ms.date: 06/08/2017
---


# ShadowFormat.Transparency Property (PowerPoint)

Returns or sets the degree of transparency of the specified fill, shadow, or line as a value between 0.0 (opaque) and 1.0 (clear). Read/write.


## Syntax

 _expression_. **Transparency**

 _expression_ A variable that represents a **ShadowFormat** object.


### Return Value

Single


## Remarks

The value of this property affects the appearance of solid-colored fills and lines only; it has no effect on the appearance of patterned lines or patterned, gradient, picture, or textured fills.


## Example

This example sets the shadow for shape three on  `myDocument` to semitransparent red. If the shape doesn't already have a shadow, this example adds one to it.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Shadow

    .Visible = True

    .ForeColor.RGB = RGB(255, 0, 0)

    .Transparency = 0.5

End With
```


## See also


#### Concepts


[ShadowFormat Object](shadowformat-object-powerpoint.md)

