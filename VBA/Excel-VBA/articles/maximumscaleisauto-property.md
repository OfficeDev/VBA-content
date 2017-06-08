---
title: MaximumScaleIsAuto Property
keywords: vbagr10.chm65572
f1_keywords:
- vbagr10.chm65572
ms.prod: excel
api_name:
- Excel.MaximumScaleIsAuto
ms.assetid: ca8115b8-0a45-0c88-5a5c-89c93d791452
ms.date: 06/08/2017
---


# MaximumScaleIsAuto Property

 **True** if Microsoft Graph calculates the maximum value for the axis. Read/write **Boolean**.


## Remarks

Setting the  **[MaximumScale](maximumscale-property.md)** property sets this property to  **False**.


## Example

This example automatically calculates the minimum scale and the maximum scale for the value axis.


```vb
With myChart.Axes(xlValue) 
 .MinimumScaleIsAuto = True 
 .MaximumScaleIsAuto = True 
End With
```


