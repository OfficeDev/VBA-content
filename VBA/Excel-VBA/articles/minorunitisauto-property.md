---
title: MinorUnitIsAuto Property
keywords: vbagr10.chm65576
f1_keywords:
- vbagr10.chm65576
ms.prod: excel
api_name:
- Excel.MinorUnitIsAuto
ms.assetid: ca6a18d5-f93f-4801-7704-4d3a25b633cb
ms.date: 06/08/2017
---


# MinorUnitIsAuto Property

 **True** if Microsoft Graph calculates minor units for the axis. Read/write **Boolean**.


## Remarks

Setting the  **[MinorUnit](minorunit-property.md)** property sets this property to  **False**.


## Example

This example automatically calculates major and minor units for the value axis.


```vb
With myChart.Axes(xlValue) 
 .MajorUnitIsAuto = True 
 .MinorUnitIsAuto = True 
End With
```


