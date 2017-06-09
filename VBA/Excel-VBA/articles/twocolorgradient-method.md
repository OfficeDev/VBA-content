---
title: TwoColorGradient Method
keywords: vbagr10.chm3077636
f1_keywords:
- vbagr10.chm3077636
ms.prod: excel
api_name:
- Excel.TwoColorGradient
ms.assetid: c42ec02c-41a2-ffc4-3d23-20a952b3de7b
ms.date: 06/08/2017
---


# TwoColorGradient Method

Sets the specified fill to a two-color gradient.

 _expression_. **TwoColorGradient**( **_Style_**,  **_Variant_**)

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

 **Style**Required 
 **MsoGradientStyle**
. Specifies the gradient style.


|MsoGradientStyle can be one of these MsoGradientStyle constants.|
| **msoGradientDiagonalDown**|
| **msoGradientDiagonalUp**|
| **msoGradientFromCenter**|
| **msoGradientFromCorner**|
| **msoGradientFromTitle**|
| **msoGradientHorizontal**|
| **msoGradientMixed**|
| **msoGradientVertical**|
 **Variant** Required **Long**. Specifies the gradient variant. Can be a value from 1 through 4, corresponding to the four variants on the  **Gradient** tab in the **Fill Effects** dialog box. If **_Style_** is **msoGradientFromCenter**, the  **_Variant_** argument can only be either 1 or 2.

## Example

This example sets the gradient, background color, and foreground color for the chart area fill on the chart.


```vb
With myChart.ChartArea.Fill 
 .Visible = True 
 .ForeColor.SchemeColor = 15 
 .BackColor.SchemeColor = 17 
 .TwoColorGradient msoGradientHorizontal, 1 
End With
```


