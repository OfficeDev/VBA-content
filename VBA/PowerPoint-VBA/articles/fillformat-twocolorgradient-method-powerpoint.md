---
title: FillFormat.TwoColorGradient Method (PowerPoint)
keywords: vbapp10.chm552008
f1_keywords:
- vbapp10.chm552008
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.TwoColorGradient
ms.assetid: 29dac3d9-366e-0fd5-0fe3-dc64fa2fc871
ms.date: 06/08/2017
---


# FillFormat.TwoColorGradient Method (PowerPoint)

Sets the specified fill to a two-color gradient.


## Syntax

 _expression_. **TwoColorGradient**( **_Style_**, **_Variant_** )

 _expression_ A variable that represents a **FillFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required|**MsoGradientStyle**|The gradient style.|
| _Variant_|Required|**Long**|The gradient variant. |

## Remarks

The  _Style_ parameter value can be one of these **MsoGradientStyle** constants.


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
The Variant parameter value can be from 1 to 4, corresponding to the four variants on the  **Gradient** sub-tab on the **Shape Fill** tab. If Style is **msoGradientFromTitle** or **msoGradientFromCenter**, this argument can be either 1 or 2.


## Example

This example adds a rectangle with a two-color gradient fill to  `myDocument` and sets the background and foreground color for the fill.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, _
        Top:=0, Width:=40, Height:=80).Fill

    .ForeColor.RGB = RGB(Red:=128, Green:=0, Blue:=0)
    .BackColor.RGB = RGB(Red:=0, Green:=170, Blue:=170)
    .TwoColorGradient Style:=msoGradientHorizontal, Variant:=1

End With
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

