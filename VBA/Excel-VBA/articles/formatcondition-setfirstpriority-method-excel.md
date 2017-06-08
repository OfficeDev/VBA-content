---
title: FormatCondition.SetFirstPriority Method (Excel)
keywords: vbaxl10.chm512091
f1_keywords:
- vbaxl10.chm512091
ms.prod: excel
api_name:
- Excel.FormatCondition.SetFirstPriority
ms.assetid: 53870387-996e-48e3-5159-7d5bb4614bcf
ms.date: 06/08/2017
---


# FormatCondition.SetFirstPriority Method (Excel)

Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.


## Syntax

 _expression_ . **SetFirstPriority**

 _expression_ A variable that represents a **FormatCondition** object.


## Remarks

When you have multiple conditional formatting rules in a worksheet, if the rule was not previously set to priority "1", this method will cause the priority of all other existing rules on the worksheet to be increased by one.


 **Note**  Priority levels for conditional formatting rules are applied on a worksheet-level basis.


## See also


#### Concepts


[FormatCondition Object](formatcondition-object-excel.md)

