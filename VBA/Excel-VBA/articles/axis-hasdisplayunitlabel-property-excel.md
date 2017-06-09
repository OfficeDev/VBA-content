---
title: Axis.HasDisplayUnitLabel Property (Excel)
keywords: vbaxl10.chm561115
f1_keywords:
- vbaxl10.chm561115
ms.prod: excel
api_name:
- Excel.Axis.HasDisplayUnitLabel
ms.assetid: 3092a94f-04ca-2d27-e21d-452b64d11f10
ms.date: 06/08/2017
---


# Axis.HasDisplayUnitLabel Property (Excel)

 **True** if the label specified by the **[DisplayUnit](axis-displayunit-property-excel.md)** or **[DisplayUnitCustom](axis-displayunitcustom-property-excel.md)** property is displayed on the specified axis. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **HasDisplayUnitLabel**

 _expression_ A variable that represents an **Axis** object.


## Example

This example sets the units on the value axis in Chart1 to increments of 500 but keeps the unit label hidden.


```vb
With Charts("Chart1").Axes(xlValue) 
 .DisplayUnit = xlCustom 
 .DisplayUnitCustom = 500 
 .AxisTitle.Caption = "Rebate Amounts" 
 .HasDisplayUnitLabel = False 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

