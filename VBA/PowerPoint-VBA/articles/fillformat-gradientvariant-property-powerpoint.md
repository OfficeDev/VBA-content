---
title: FillFormat.GradientVariant Property (PowerPoint)
keywords: vbapp10.chm552016
f1_keywords:
- vbapp10.chm552016
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.GradientVariant
ms.assetid: 32a8a1fd-84aa-fbee-35c5-5bd83b0790c6
ms.date: 06/08/2017
---


# FillFormat.GradientVariant Property (PowerPoint)

Returns the gradient variant for the specified fill as an integer value from 1 to 4 for most gradient fills. Read-only.


## Syntax

 _expression_. **GradientVariant**

 _expression_ A variable that represents a **FillFormat** object.


### Return Value

Long


## Remarks

 If the gradient style is **msoGradientFromTitle** or **msoGradientFromCenter**, this property returns either 1 or 2.

The values for this property correspond to the gradient variants (numbered from left to right and from top to bottom) on the  **Gradient** subtab in the **Shape Fill** tab. **Long**.

This property is read-only. Use the [OneColorGradient](fillformat-onecolorgradient-method-powerpoint.md), [PresetGradient](fillformat-presetgradient-method-powerpoint.md), or  **[TwoColorGradient](fillformat-twocolorgradient-method-powerpoint.md)** method to set the gradient variant for the fill.


## Example

This example adds a rectangle to  `myDocument` and sets its fill gradient variant to match that of the shape named "rect1." For the example to work, rect1 must have a gradient fill.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes

    gradVar1 = .Item("rect1").Fill.GradientVariant

    With .AddShape(msoShapeRectangle, 0, 0, 40, 80).Fill

        .ForeColor.RGB = RGB(128, 0, 0)

        .OneColorGradient msoGradientHorizontal, gradVar1, 1

    End With

End With
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

