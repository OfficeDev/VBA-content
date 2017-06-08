---
title: Axis.MinimumScale Property (Excel)
keywords: vbaxl10.chm561090
f1_keywords:
- vbaxl10.chm561090
ms.prod: excel
api_name:
- Excel.Axis.MinimumScale
ms.assetid: 31cfa07e-24a6-666f-7bb0-6bb5c139d4d9
ms.date: 06/08/2017
---


# Axis.MinimumScale Property (Excel)

Returns or sets the minimum value on the value axis. Read/write  **Double** .


## Syntax

 _expression_ . **MinimumScale**

 _expression_ A variable that represents an **Axis** object.


## Remarks

Setting this property sets the  **[MinimumScaleIsAuto](axis-minimumscaleisauto-property-excel.md)** property to **False** .


## Example

This example sets the minimum and maximum values for the value axis in Chart1.


```vb
With Charts("Chart1").Axes(xlValue) 
 .MinimumScale = 10 
 .MaximumScale = 120 
End With
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

