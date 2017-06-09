---
title: FillFormat.PresetGradient Method (Publisher)
keywords: vbapb10.chm2359315
f1_keywords:
- vbapb10.chm2359315
ms.prod: publisher
api_name:
- Publisher.FillFormat.PresetGradient
ms.assetid: d97c4ce8-5cef-6f53-d0c8-8bcf9ab8bb80
ms.date: 06/08/2017
---


# FillFormat.PresetGradient Method (Publisher)

Sets the specified fill to a preset gradient.


## Syntax

 _expression_. **PresetGradient**( **_Style_**,  **_Variant_**,  **_PresetGradientType_**)

 _expression_A variable that represents a  **FillFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Style|Required| **MsoGradientStyle**|The style of the gradient.|
|Variant|Required| **Long**|The gradient variant. Can be a value from 1 to 4, corresponding to the four variants on the  **Gradient** tab in the **Fill Effects** dialog box. If Style is **msoGradientFromTitle** or **msoGradientFromCenter**, this argument can be either 1 or 2.|
|PresetGradientType|Required| **MsoPresetGradientType**|The gradient type.|

## Remarks

The Style parameter can be one of the  **MsoPresetGradientStyle** constants declared in the Microsoft Office type library and shown in the following table.



| **msoGradientDiagonalDown**|
| **msoGradientDiagonalUp**|
| **msoGradientFromCenter**|
| **msoGradientFromCorner**|
| **msoGradientFromTitle**|
| **msoGradientHorizontal**|
| **msoGradientVertical**|
The PresetGradientType parameter can be one of the  **MsoPresetGradientType** constants declared in the Microsoft Office type library and shown in the following table.



| **msoGradientBrass**|
| **msoGradientCalmWater**|
| **msoGradientChrome**|
| **msoGradientChromeII**|
| **msoGradientDaybreak**|
| **msoGradientDesert**|
| **msoGradientEarlySunset**|
| **msoGradientFire**|
| **msoGradientFog**|
| **msoGradientGold**|
| **msoGradientGoldII**|
| **msoGradientHorizon**|
| **msoGradientLateSunset**|
| **msoGradientMahogany**|
| **msoGradientMoss**|
| **msoGradientNightfall**|
| **msoGradientOcean**|
| **msoGradientParchment**|
| **msoGradientPeacock**|
| **msoGradientRainbow**|
| **msoGradientRainbowII**|
| **msoGradientSapphire**|
| **msoGradientSilver**|
| **msoGradientWheat**|

## Example

This example adds a rectangle with a preset gradient fill to the active publication.


```vb
ActiveDocument.Pages(1).Shapes _ 
 .AddShape(msoShapeRectangle, 90, 90, 140, 80) _ 
 .Fill.PresetGradient Style:=msoGradientHorizontal, _ 
 Variant:=1, PresetGradientType:=msoGradientBrass 

```


