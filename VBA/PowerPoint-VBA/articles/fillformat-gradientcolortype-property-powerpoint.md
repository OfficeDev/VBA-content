---
title: FillFormat.GradientColorType Property (PowerPoint)
keywords: vbapp10.chm552013
f1_keywords:
- vbapp10.chm552013
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.GradientColorType
ms.assetid: 90224ee2-80f9-480b-bd1b-678035ded3ef
ms.date: 06/08/2017
---


# FillFormat.GradientColorType Property (PowerPoint)

Returns the gradient color type for the specified fill. Read-only.


## Syntax

 _expression_. **GradientColorType**

 _expression_ A variable that represents a **FillFormat** object.


### Return Value

MsoGradientColorType


## Remarks

Use the [OneColorGradient](fillformat-onecolorgradient-method-powerpoint.md), [PresetGradient](fillformat-presetgradient-method-powerpoint.md), or  **[TwoColorGradient](fillformat-twocolorgradient-method-powerpoint.md)** method to set the gradient type for the fill.

The value returned by the  **GradientColorType** property can be one of these **MsoGradientColorType** constants.


||
|:-----|
|**msoGradientColorMixed**|
|**msoGradientOneColor**|
|**msoGradientPresetColors**|
|**msoGradientTwoColors**|

## Example

This example changes the fill for all shapes in  `myDocument` that have a two-color gradient fill to a preset gradient fill.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes
    With s.Fill
        If .GradientColorType = msoGradientTwoColors Then
            .PresetGradient msoGradientHorizontal, _
                1, msoGradientBrass
        End If
    End With
Next
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

