---
title: Axis.MaximumScaleIsAuto Property (Excel)
keywords: vbaxl10.chm561089
f1_keywords:
- vbaxl10.chm561089
ms.prod: excel
api_name:
- Excel.Axis.MaximumScaleIsAuto
ms.assetid: c0e0f4b6-5d1c-5acb-2e7a-8722e10cd2bc
ms.date: 06/08/2017
---


# Axis.MaximumScaleIsAuto Property (Excel)

 **True** if Microsoft Excel calculates the maximum value for the value axis. Read/write **Boolean** .


## Syntax

 _expression_ . **MaximumScaleIsAuto**

 _expression_ A variable that represents an **Axis** object.


## Remarks

Setting the  **[MaximumScale](axis-maximumscale-property-excel.md)** property sets this property to **False** .


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

