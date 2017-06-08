---
title: IconSetCondition.SetFirstPriority Method (Excel)
keywords: vbaxl10.chm812080
f1_keywords:
- vbaxl10.chm812080
ms.prod: excel
api_name:
- Excel.IconSetCondition.SetFirstPriority
ms.assetid: 9d37baef-3e0d-95fa-a251-d60f20830625
ms.date: 06/08/2017
---


# IconSetCondition.SetFirstPriority Method (Excel)

Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.


## Syntax

 _expression_ . **SetFirstPriority**

 _expression_ A variable that represents an **IconSetCondition** object.


## Remarks

When you have multiple conditional formatting rules in a worksheet, if the rule was not previously set to priority "1", this method will cause the priority of all other existing rules on the worksheet to be increased by one.


 **Note**  Priority levels for conditional formatting rules are applied on a worksheet-level basis.


## See also


#### Concepts


[IconSetCondition Object](iconsetcondition-object-excel.md)

