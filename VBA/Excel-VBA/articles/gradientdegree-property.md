---
title: GradientDegree Property
keywords: vbagr10.chm67177
f1_keywords:
- vbagr10.chm67177
ms.prod: excel
api_name:
- Excel.GradientDegree
ms.assetid: 6f325dc0-5f6c-7a55-52fa-55eeb15ccfe6
ms.date: 06/08/2017
---


# GradientDegree Property

Returns the gradient degree of the specified one-color shaded fill as a floating-point value from 0.0 (dark) through 1.0 (light). Read-only Single.

This property is read-only. Use the OneColorGradient method to set the gradient degree for the fill.

 _expression_. **GradientDegree**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.

## Example

This example sets the chart's fill format so that its gradient degree is at least 0.3.


```vb
With myChart.ChartArea.Fill 
 If .Type = msoFillGradient Then 
 If .GradientColorType = msoGradientOneColor Then 
 If .GradientDegree < 0.3 Then 
 .OneColorGradient .GradientStyle, _ 
 .GradientVariant, 0.3 
 End If 
 End If 
 End If 
End With 

```


