---
title: FormatCondition.SetLastPriority Method (Excel)
keywords: vbaxl10.chm512092
f1_keywords:
- vbaxl10.chm512092
ms.prod: excel
api_name:
- Excel.FormatCondition.SetLastPriority
ms.assetid: fd6263a1-e67f-f4e8-2423-1601f73bdd5c
ms.date: 06/08/2017
---


# FormatCondition.SetLastPriority Method (Excel)

Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.


## Syntax

 _expression_ . **SetLastPriority**

 _expression_ A variable that represents a **FormatCondition** object.


## Remarks

The actual value of the priority will be equal to the total number of conditional formatting rules on the worksheet. When you have multiple conditional formatting rules in a worksheet, this method will cause the priority of rules that had a priority value greater than this rule to be decreased by one.


 **Note**  Priority levels for conditional formatting rules are applied on a worksheet-level basis.


## See also


#### Concepts


[FormatCondition Object](formatcondition-object-excel.md)

