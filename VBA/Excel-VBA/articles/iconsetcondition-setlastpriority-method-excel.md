---
title: IconSetCondition.SetLastPriority Method (Excel)
keywords: vbaxl10.chm812081
f1_keywords:
- vbaxl10.chm812081
ms.prod: excel
api_name:
- Excel.IconSetCondition.SetLastPriority
ms.assetid: b1003681-b5ac-85ab-dd9c-8a13685694d6
ms.date: 06/08/2017
---


# IconSetCondition.SetLastPriority Method (Excel)

Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.


## Syntax

 _expression_ . **SetLastPriority**

 _expression_ A variable that represents an **IconSetCondition** object.


## Remarks

The actual value of the priority will be equal to the total number of conditional formatting rules on the worksheet. When you have multiple conditional formatting rules in a worksheet, this method will cause the priority of rules that had a priority value greater than this rule to be decreased by one.


 **Note**  Priority levels for conditional formatting rules are applied on a worksheet-level basis.


## See also


#### Concepts


[IconSetCondition Object](iconsetcondition-object-excel.md)

