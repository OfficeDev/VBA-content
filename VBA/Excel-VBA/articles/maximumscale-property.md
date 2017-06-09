---
title: MaximumScale Property
keywords: vbagr10.chm5207676
f1_keywords:
- vbagr10.chm5207676
ms.prod: excel
api_name:
- Excel.MaximumScale
ms.assetid: 1fd6633e-7782-78d0-ba24-9c3d46f85471
ms.date: 06/08/2017
---


# MaximumScale Property

Returns or sets the maximum value on the axis. Read/write  **Double**.


## Remarks

Setting this property sets the  **[MaximumScaleIsAuto](maximumscaleisauto-property.md)** property to  **False**.


## Example

This example sets the minimum and maximum values for the value axis.


```vb
With myChart.Axes(xlValue) 
 .MinimumScale = 10 
 .MaximumScale = 120 
End With
```


