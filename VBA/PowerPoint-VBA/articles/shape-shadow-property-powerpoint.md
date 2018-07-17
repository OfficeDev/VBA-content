---
title: Shape.Shadow Property (PowerPoint)
keywords: vbapp10.chm547033
f1_keywords:
- vbapp10.chm547033
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Shadow
ms.assetid: 832b8e62-4fc5-1f4b-74c7-cc0e63a12699
ms.date: 06/08/2017
---


# Shape.Shadow Property (PowerPoint)

Returns a  **[ShadowFormat](shadowformat-object-powerpoint.md)** object that contains shadow formatting properties for the specified shape. Read-only.


## Syntax

 _expression_. **Shadow**

 _expression_ A variable that represents a **Shape** object.


## Example

This example adds a shadowed rectangle to slide one in the active presentation. The blue, embossed shadow is offset 3 points to the right of and 2 points down from the rectangle.


```vb
Set myShap = Application.ActivePresentation.Slides(1).Shapes

With myShap.AddShape(msoShapeRectangle, 10, 10, 150, 90).Shadow

    .Type = msoShadow17

    .ForeColor.RGB = RGB(0, 0, 128)

    .OffsetX = 3

    .OffsetY = 2

End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

