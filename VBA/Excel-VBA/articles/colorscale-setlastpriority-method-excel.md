---
title: ColorScale.SetLastPriority Method (Excel)
keywords: vbaxl10.chm806079
f1_keywords:
- vbaxl10.chm806079
ms.prod: excel
api_name:
- Excel.ColorScale.SetLastPriority
ms.assetid: 01c64e4d-98e8-3647-5e06-23fd1000757b
ms.date: 06/08/2017
---


# ColorScale.SetLastPriority Method (Excel)

Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.


## Syntax

 _expression_ . **SetLastPriority**

 _expression_ A variable that represents a **ColorScale** object.


## Remarks

The actual value of the priority will be equal to the total number of conditional formatting rules on the worksheet. When you have multiple conditional formatting rules in a worksheet, this method will cause the priority of rules that had a priority value greater than this rule to be decreased by one.


 **Note**  Priority levels for conditional formatting rules are applied on a worksheet-level basis.


## See also


#### Concepts


[ColorScale Object](colorscale-object-excel.md)

