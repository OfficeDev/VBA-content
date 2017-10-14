---
title: ShadowFormat.OffsetY Property (PowerPoint)
keywords: vbapp10.chm554007
f1_keywords:
- vbapp10.chm554007
ms.prod: powerpoint
api_name:
- PowerPoint.ShadowFormat.OffsetY
ms.assetid: 286f5d2a-5601-f3d7-0ace-fc01f168224d
ms.date: 06/08/2017
---


# ShadowFormat.OffsetY Property (PowerPoint)

Returns or sets the vertical offset of the shadow from the specified shape, in points. Read/write.


## Syntax

 _expression_. **OffsetY**

 _expression_ A variable that represents an **ShadowFormat** object.


### Return Value

Single


## Remarks

 A positive value offsets the shadow to the right of the shape; a negative value offsets it to the left.

If you want to nudge a shadow horizontally or vertically from its current position without having to specify an absolute position, use the  **[IncrementOffsetX](shadowformat-incrementoffsetx-method-powerpoint.md)** method or the **[IncrementOffsetY](shadowformat-incrementoffsety-method-powerpoint.md)** method.


## Example

This example sets the horizontal and vertical offsets of the shadow for shape three on  `myDocument`. The shadow is offset 5 points to the right of the shape and 3 points above it. If the shape doesn't already have a shadow, this example adds one to it.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Shadow

    .Visible = True

    .OffsetX = 5

    .OffsetY = -3

End With
```


## See also


#### Concepts


[ShadowFormat Object](shadowformat-object-powerpoint.md)

