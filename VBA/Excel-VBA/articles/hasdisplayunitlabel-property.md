---
title: HasDisplayUnitLabel Property
keywords: vbagr10.chm5241529
f1_keywords:
- vbagr10.chm5241529
ms.prod: excel
api_name:
- Excel.HasDisplayUnitLabel
ms.assetid: 5093286f-53ff-3c56-d047-7b6a92d2b7d6
ms.date: 06/08/2017
---


# HasDisplayUnitLabel Property

 **True** if the label specified by the **[DisplayUnit](displayunit-property.md)** or  **[DisplayUnitCustom](displayunitcustom-property.md)** property is displayed on the value axis.  **False** if no units are displayed. The default value is **True**. Read/write  **Boolean**.


## Example

This example sets the units on the value axis in myChart to increments of 500 but hides the unit label itself.


```vb
With myChart.Axes(xlValue) 
 .DisplayUnit = xlCustom 
 .DisplayUnitCustom = 500 
 .AxisTitle.Caption = "Rebate Amounts" 
 .HasDisplayUnitLabel = False 
End With
```


