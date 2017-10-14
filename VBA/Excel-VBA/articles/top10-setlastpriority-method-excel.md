---
title: Top10.SetLastPriority Method (Excel)
keywords: vbaxl10.chm822085
f1_keywords:
- vbaxl10.chm822085
ms.prod: excel
api_name:
- Excel.Top10.SetLastPriority
ms.assetid: 878cbcd5-47c9-64f8-d864-cfe279dec513
ms.date: 06/08/2017
---


# Top10.SetLastPriority Method (Excel)

Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.


## Syntax

 _expression_ . **SetLastPriority**

 _expression_ A variable that represents a **Top10** object.


## Remarks

The actual value of the priority will be equal to the total number of conditional formatting rules on the worksheet. When you have multiple conditional formatting rules in a worksheet, this method will cause the priority of rules that had a priority value greater than this rule to be decreased by one.


 **Note**  Priority levels for conditional formatting rules are applied on a worksheet-level basis.


## See also


#### Concepts


[Top10 Object](top10-object-excel.md)

