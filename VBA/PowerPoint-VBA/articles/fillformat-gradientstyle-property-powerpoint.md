---
title: FillFormat.GradientStyle Property (PowerPoint)
keywords: vbapp10.chm552015
f1_keywords:
- vbapp10.chm552015
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.GradientStyle
ms.assetid: dca37bf2-1219-d815-7584-97a8665e3420
ms.date: 06/08/2017
---


# FillFormat.GradientStyle Property (PowerPoint)

Returns the gradient style for the specified fill. Read-only.


## Syntax

 _expression_. **GradientStyle**

 _expression_ A variable that represents a **FillFormat** object.


### Return Value

MsoGradientStyle


## Remarks

Use the [OneColorGradient](fillformat-onecolorgradient-method-powerpoint.md), [PresetGradient](fillformat-presetgradient-method-powerpoint.md), or  **[TwoColorGradient](fillformat-twocolorgradient-method-powerpoint.md)** method to set the gradient style for the fill. Attempting to return this property for a fill that doesn't have a gradient generates an error. Use the **[Type](filtereffect-type-property-powerpoint.md)** property to determine whether the fill has a gradient.

The value returned by the  **GradientStyle** property can be one of these **MsoGradientStyle** constants.


||
|:-----|
|**msoGradientDiagonalDown**|
|**msoGradientDiagonalUp**|
|**msoGradientFromCenter**|
|**msoGradientFromCorner**|
|**msoGradientFromTitle**|
|**msoGradientHorizontal**|
|**msoGradientMixed**|
|**msoGradientVertical**|

## Example

This example adds a rectangle to  `myDocument` and sets its fill gradient style to match that of the shape named "rect1." For the example to work, rect1 must have a gradient fill.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    gradStyle1 = .Item("rect1").Fill.GradientStyle

    With .AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill

        .ForeColor.RGB = RGB(128, 0, 0)

        .OneColorGradient gradStyle1, 1, 1

    End With

End With
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

