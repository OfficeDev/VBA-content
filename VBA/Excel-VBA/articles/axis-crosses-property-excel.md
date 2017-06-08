---
title: Axis.Crosses Property (Excel)
keywords: vbaxl10.chm561078
f1_keywords:
- vbaxl10.chm561078
ms.prod: excel
api_name:
- Excel.Axis.Crosses
ms.assetid: 571e256d-b711-e3cd-f0f2-c53e86375e6f
ms.date: 06/08/2017
---


# Axis.Crosses Property (Excel)

Returns or sets the point on the specified axis where the other axis crosses. Read/write  **Long** .


## Syntax

 _expression_ . **Crosses**

 _expression_ A variable that represents an **Axis** object.


## Remarks

Can be one of the  **XlAxisCrosses** constants listed in the following table.



|**Constant**|**Meaning**|
|:-----|:-----|
| **xlAxisCrossesAutomatic**|Microsoft Excel sets the axis crossing point.|
| **xlMinimum**|The axis crosses at the minimum value.|
| **xlMaximum**|The axis crosses at the maximum value.|
| **xlAxisCrossesCustom**|The  **[CrossesAt](axis-crossesat-property-excel.md)** property specifies the axis crossing point.|
This property isn't available for radar charts. For 3-D charts, this property can only be applied to the value axis and indicates where the plane defined by the category axes crosses the value axis.

This property can be used for both category and value axes. On the category axis,  **xlMinimum** sets the value axis to cross at the first category, and **xlMaximum** sets the value axis to cross at the last category.

Note that  **xlMinimum** and **xlMaximum** can have different meanings, depending on the axis.


## Example

This example sets the value axis in Chart1 to cross the category axis at the maximum x value.


```vb
Charts("Chart1").Axes(xlCategory).Crosses = xlMaximum
```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

