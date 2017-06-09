---
title: FillFormat.PresetGradient Method (PowerPoint)
keywords: vbapp10.chm552005
f1_keywords:
- vbapp10.chm552005
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.PresetGradient
ms.assetid: 6aa304c7-a2ee-ceea-f956-404538bebc43
ms.date: 06/08/2017
---


# FillFormat.PresetGradient Method (PowerPoint)

Sets the specified fill to a preset gradient.


## Syntax

 _expression_. **PresetGradient**( **_Style_**, **_Variant_**, **_PresetGradientType_** )

 _expression_ A variable that represents a **FillFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Required|**MsoGradientStyle**|The gradient style.|
| _Variant_|Required|**Integer**|The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the  **Gradient** subtab on the **Shape Fill** tab. If Style is **msoGradientFromTitle** or **msoGradientFromCenter**, this argument can be either 1 or 2.|
| _PresetGradientType_|Required|**MsoPresetGradientType**|The gradient type.|

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
The  _PresetGradientType_ parameter value can be one of these **MsoPresetGradientType** constants.


||
|:-----|
|**msoGradientBrass**|
|**msoGradientCalmWater**|
|**msoGradientChrome**|
|**msoGradientChromeII**|
|**msoGradientDaybreak**|
|**msoGradientDesert**|
|**msoGradientEarlySunset**|
|**msoGradientFire**|
|**msoGradientFog**|
|**msoGradientGold**|
|**msoGradientGoldII**|
|**msoGradientHorizon**|
|**msoGradientLateSunset**|
|**msoGradientMahogany**|
|**msoGradientMoss**|
|**msoGradientNightfall**|
|**msoGradientOcean**|
|**msoGradientParchment**|
|**msoGradientPeacock**|
|**msoGradientRainbow**|
|**msoGradientRainbowII**|
|**msoGradientSapphire**|
|**msoGradientSilver**|
|**msoGradientWheat**|
|**msoPresetGradientMixed**|

## Example

This example adds a rectangle with a preset gradient fill to  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 140, 80).Fill.PresetGradient msoGradientHorizontal, 1, msoGradientBrass
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

