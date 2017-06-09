---
title: PresetGradient Method
keywords: vbagr10.chm67172
f1_keywords:
- vbagr10.chm67172
ms.prod: excel
api_name:
- Excel.PresetGradient
ms.assetid: db282722-c2ad-b504-62b3-326814fd8ca0
ms.date: 06/08/2017
---


# PresetGradient Method

Sets the specified fill to a preset gradient.

 _expression_. **PresetGradient**( **_Style_**,  **_Variant_**,  **_PresetGradientType_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Style**Required 
 **MsoGradientStyle**
. The gradient style for the specified fill.


|MsoGradientStyle can be one of these MsoGradientStyle constants.|
| **msoGradientDiagonalDown**|
| **msoGradientDiagonalUp**|
| **msoGradientFromCenter**|
| **msoGradientFromCorner**|
| **msoGradientFromTitle**|
| **msoGradientHorizontal**|
| **msoGradientMixed**|
| **msoGradientVertical**|
 **Variant**Required  **Long**. The gradient variant for the specified fill. Can be a value from 1 through 4, corresponding to the four variants listed on the  **Gradient** tab in the **Fill Effects** dialog box. If **_Style_** is **msoGradientFromCenter**, the  **_Variant_** argument can only be 1 or 2.
 **PresetGradientType**Required 
 **MsoPresetGradientType**
. The gradient type for the specified fill.


|MsoPresetGradientType can be one of these MsoPresetGradientType constants.|
| **msoGradientBrass**|
| **msoGradientChrome**|
| **msoGradientDaybreak**|
| **msoGradientEarlySunset**|
| **msoGradientFog**|
| **msoGradientGoldII**|
| **msoGradientLateSunset**|
| **msoGradientMoss**|
| **msoGradientOcean**|
| **msoGradientPeacock**|
| **msoGradientRainbowII**|
| **msoGradientSilver**|
| **msoGradientWheat**|
| **msoPresetGradientMixed**|
| **msoGradientCalmWater**|
| **msoGradientChromeII**|
| **msoGradientDesert**|
| **msoGradientFire**|
| **msoGradientGold**|
| **msoGradientHorizon**|
| **msoGradientMahogany**|
| **msoGradientNightfall**|
| **msoGradientParchment**|
| **msoGradientRainbow**|
| **msoGradientSapphire**|

## Example

This example sets the chart's fill format to the preset brass color.


```vb
With myChart.ChartArea.Fill 
 .Visible = True 
 .PresetGradient msoGradientDiagonalDown, 3, msoGradientBrass 
End With
```


