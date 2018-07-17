---
title: Databar.SetLastPriority Method (Excel)
keywords: vbaxl10.chm810085
f1_keywords:
- vbaxl10.chm810085
ms.prod: excel
api_name:
- Excel.Databar.SetLastPriority
ms.assetid: 985b1225-6816-fe3b-e973-5fd90aa1fe47
ms.date: 06/08/2017
---


# Databar.SetLastPriority Method (Excel)

Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.


## Syntax

 _expression_ . **SetLastPriority**

 _expression_ A variable that represents a **Databar** object.


## Remarks

The actual value of the priority will be equal to the total number of conditional formatting rules on the worksheet. When you have multiple conditional formatting rules in a worksheet, this method will cause the priority of rules that had a priority value greater than this rule to be decreased by one.


 **Note**  Priority levels for conditional formatting rules are applied on a worksheet-level basis.


## See also


#### Concepts


[Databar Object](databar-object-excel.md)

