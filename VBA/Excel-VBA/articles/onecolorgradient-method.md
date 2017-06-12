---
title: OneColorGradient Method
keywords: vbagr10.chm67157
f1_keywords:
- vbagr10.chm67157
ms.prod: excel
api_name:
- Excel.OneColorGradient
ms.assetid: 7e572d28-2905-2c6b-5e62-1f763bba7f89
ms.date: 06/08/2017
---


# OneColorGradient Method

Sets the specified fill to a one-color gradient.

 _expression_. **OneColorGradient**( **_Style_**,  **_Variant_**,  **_Degree_**)

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
 **Degree**Required  **Single**. The gradient degree for the specified fill. Can be a value from 0.0 (dark) through 1.0 (light).

## Example

This example sets the chart's fill format.


```vb
With myChart.ChartArea.Fill 
 If .Type = msoFillGradient Then 
 If .GradientColorType = msoGradientOneColor Then 
 .OneColorGradient Style:=msoGradientFromCorner, _ 
 Variant:=1, Degree:=0.3 
 End If 
 End If 
End With
```


