---
title: MajorUnitIsAuto Property
keywords: vbagr10.chm5207645
f1_keywords:
- vbagr10.chm5207645
ms.prod: excel
api_name:
- Excel.MajorUnitIsAuto
ms.assetid: 6eda8012-2ef3-d23b-bace-e2695a5e80f5
ms.date: 06/08/2017
---


# MajorUnitIsAuto Property

 **True** if Microsoft Graph calculates the major units for the axis. Read/write **Boolean**.


## Remarks

Setting the  **[MajorUnit](majorunit-property.md)** property sets this property to  **False**.


## Example

This example automatically sets the major and minor units for the value axis.


```vb
With myChart.Axes(xlValue) 
 .MajorUnitIsAuto = True 
 .MinorUnitIsAuto = True 
End With
```


