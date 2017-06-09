---
title: Axis.MajorUnitIsAuto Property (Excel)
keywords: vbaxl10.chm561087
f1_keywords:
- vbaxl10.chm561087
ms.prod: excel
api_name:
- Excel.Axis.MajorUnitIsAuto
ms.assetid: bec8cc5a-c4c9-7d59-bf0d-ae88b9891182
ms.date: 06/08/2017
---


# Axis.MajorUnitIsAuto Property (Excel)

 **True** if Microsoft Excel calculates the major units for the value axis. Read/write **Boolean** .


## Syntax

 _expression_ . **MajorUnitIsAuto**

 _expression_ A variable that represents an **Axis** object.


## Remarks

Setting the  **[MajorUnit](axis-majorunit-property-excel.md)** property sets this property to **False** .


## Example

This example automatically sets the major and minor units for the value axis in Chart1.


```vb
With Charts("Chart1").Axes(xlValue) 
 .MajorUnitIsAuto = True 
 .MinorUnitIsAuto = True 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

