---
title: Axis.MinimumScaleIsAuto Property (Excel)
keywords: vbaxl10.chm561091
f1_keywords:
- vbaxl10.chm561091
ms.prod: excel
api_name:
- Excel.Axis.MinimumScaleIsAuto
ms.assetid: 93767cb3-c71e-b191-2f07-7ca091498023
ms.date: 06/08/2017
---


# Axis.MinimumScaleIsAuto Property (Excel)

 **True** if Microsoft Excel calculates the minimum value for the value axis. Read/write **Boolean** .


## Syntax

 _expression_ . **MinimumScaleIsAuto**

 _expression_ A variable that represents an **Axis** object.


## Remarks

Setting the  **[MinimumScale](axis-minimumscale-property-excel.md)** property sets this property to **False** .


## Example

This example automatically calculates the minimum scale and the maximum scale for the value axis in Chart1.


```vb
With Charts("Chart1").Axes(xlValue) 
 .MinimumScaleIsAuto = True 
 .MaximumScaleIsAuto = True 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

