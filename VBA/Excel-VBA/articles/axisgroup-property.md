---
title: AxisGroup Property
keywords: vbagr10.chm65583
f1_keywords:
- vbagr10.chm65583
ms.prod: excel
api_name:
- Excel.AxisGroup
ms.assetid: 453bc2f6-ca27-1b7c-8dc4-8a902c9445be
ms.date: 06/08/2017
---


# AxisGroup Property

AxisGroup property as it applies to the  **ChartGroup** and **Series** objects.

Returns the group for the specified chart group or series. Read/write XlAxisGroup .


|XlAxisGroup can be one of these XlAxisGroup constants.|
| **xlPrimary**|
| **xlSecondary**|
 _expression_. **AxisGroup**
 _expression_ Required. An expression that returns one of the above objects.
AxisGroup property as it applies to the  **Axis** object.
Returns the group for the specified axis. Read-only XlAxisGroup .


|XlAxisGroup can be one of these XlAxisGroup constants.|
| **xlPrimary**|
| **xlSecondary**|
 _expression_. **AxisGroup**
 _expression_ Required. An expression that returns one of the above objects.

## Remarks

For 3-D charts, only  **xlPrimary** is valid.


## Example

This example deletes the value axis if it's in the secondary group.


```vb
With myChart.Axes(xlValue) 
 If .AxisGroup = xlSecondary Then .Delete 
End With
```


