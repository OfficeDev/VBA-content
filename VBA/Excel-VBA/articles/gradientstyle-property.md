---
title: GradientStyle Property
keywords: vbagr10.chm3077038
f1_keywords:
- vbagr10.chm3077038
ms.prod: excel
api_name:
- Excel.GradientStyle
ms.assetid: 042a271c-165c-ba10-9702-7db744654760
ms.date: 06/08/2017
---


# GradientStyle Property

Returns the gradient style for the specified fill. Read-only MsoGradientStyle .



|MsoGradientStyle can be one of these MsoGradientStyle constants.|
| **msoGradientDiagonalDown**|
| **msoGradientDiagonalUp**|
| **msoGradientFromCenter**|
| **msoGradientFromCorner**|
| **msoGradientFromTitle**|
| **msoGradientHorizontal**|
| **msoGradientMixed**|
| **msoGradientVertical**This property is read-only. Use the  **OneColorGradient** or **TwoColorGradient** method to set the gradient style for the fill.|

 _expression_. **GradientStyle**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets the chart's fill format so that its gradient style is changed to  **msoGradientDiagonalUp** if it was originally **msoGradientDiagonalDown**.


```vb
With myChart.ChartArea.Fill 
 If .Type = msoFillGradient Then 
 If .GradientColorType = msoGradientOneColor Then 
 If .GradientStyle = msoGradientDiagonalDown Then 
 .OneColorGradient msoGradientDiagonalUp, _ 
 .GradientVariant, .GradientDegree 
 End If 
 End If 
 End If 
End With
```


