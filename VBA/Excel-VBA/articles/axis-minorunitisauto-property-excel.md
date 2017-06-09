---
title: Axis.MinorUnitIsAuto Property (Excel)
keywords: vbaxl10.chm561095
f1_keywords:
- vbaxl10.chm561095
ms.prod: excel
api_name:
- Excel.Axis.MinorUnitIsAuto
ms.assetid: fff34170-5073-9053-4059-83d29ba9d399
ms.date: 06/08/2017
---


# Axis.MinorUnitIsAuto Property (Excel)

 **True** if Microsoft Excel calculates minor units for the value axis. Read/write **Boolean** .


## Syntax

 _expression_ . **MinorUnitIsAuto**

 _expression_ A variable that represents an **Axis** object.


## Remarks

Setting the  **[MinorUnit](axis-minorunit-property-excel.md)** property sets this property to **False** .


## Example

This example automatically calculates major and minor units for the value axis in Chart1.


```vb
With Charts("Chart1").Axes(xlValue) 
 .MajorUnitIsAuto = True 
 .MinorUnitIsAuto = True 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

