---
title: Databar.PercentMin Property (Excel)
keywords: vbaxl10.chm810078
f1_keywords:
- vbaxl10.chm810078
ms.prod: excel
api_name:
- Excel.Databar.PercentMin
ms.assetid: bd8670f9-ae0b-3a1c-5b14-84cc00638b6e
ms.date: 06/08/2017
---


# Databar.PercentMin Property (Excel)

Returns or sets a  **Long** value that specifies the length of the shortest data bar as a percentage of cell width.


## Syntax

 _expression_ . **PercentMin**

 _expression_ A variable that represents a **[Databar](databar-object-excel.md)** object.


## Remarks

The value must be a whole number between 0 and 100. The default value is 0.

The effect of the  **PercentMin** property varies depending on the setting of the **[AxisPosition](databar-axisposition-property-excel.md)** property of the **Databar** object. When the **AxisPosition** property is **xlDataBarAxisAutomatic** and the range contains both positive and negative values, the minimum length of a positive or negative bar is specified by the **PercentMin** property, and the axis is displayed using automatic centering rules. When the **AxisPosition** property is **xlDataBarAxisMidpoint** , the minimum length of a positive or negative bar is specified by the **PercentMin** property, and the axis is centered in the middle of the cell. When the **AxisPosition** property is **xlDataBarAxisNone** , the length of the shortest data bar is always the percentage of cell width specified by the **PercentMin** property.


## See also


#### Concepts


[Databar Object](databar-object-excel.md)

