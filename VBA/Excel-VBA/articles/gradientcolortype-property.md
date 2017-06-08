---
title: GradientColorType Property
keywords: vbagr10.chm3077036
f1_keywords:
- vbagr10.chm3077036
ms.prod: excel
api_name:
- Excel.GradientColorType
ms.assetid: 78a2bd69-e8a5-1c43-4c75-9715de4202c0
ms.date: 06/08/2017
---


# GradientColorType Property

Returns the gradient color type for the specified fill. Read-only MsoGradientColorType .



|MsoGradientColorType can be one of these MsoGradientColorType constants.|
| **msoGradientColorMixed**|
| **msoGradientOneColor**|
| **msoGradientPresetColors**|
| **msoGradientTwoColors**|

 _expression_. **GradientColorType**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets the fill format for the chart if its chart area has a one-color gradient fill.


```vb
With myChart.ChartArea.Fill 
 If .Type = msoFillGradient Then 
 If .GradientColorType = msoGradientOneColor Then 
 .OneColorGradient Style:= msoGradientFromCorner, _ 
 Variant:= 1, Degree:= 0.3 
 End If 
 End If 
End With
```


