---
title: MinimumScaleIsAuto Property
keywords: vbagr10.chm5207691
f1_keywords:
- vbagr10.chm5207691
ms.prod: excel
api_name:
- Excel.MinimumScaleIsAuto
ms.assetid: 95ed7a2b-efda-b05a-da2e-789a166a97c8
ms.date: 06/08/2017
---


# MinimumScaleIsAuto Property

 **True** if Microsoft Graph calculates the minimum value for the axis. Read/write **Boolean**.


## Remarks

Setting the  **[MinimumScale](minimumscale-property.md)** property sets this property to  **False**.


## Example

This example automatically calculates the minimum scale and the maximum scale for the value axis.


```vb
With myChart.Axes(xlValue) 
 .MinimumScaleIsAuto = True 
 .MaximumScaleIsAuto = True 
End With
```


