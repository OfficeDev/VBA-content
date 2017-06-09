---
title: Databar.SetFirstPriority Method (Excel)
keywords: vbaxl10.chm810084
f1_keywords:
- vbaxl10.chm810084
ms.prod: excel
api_name:
- Excel.Databar.SetFirstPriority
ms.assetid: 73ec6aa8-dc0d-7f80-0975-fdf75bd9a0a2
ms.date: 06/08/2017
---


# Databar.SetFirstPriority Method (Excel)

Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.


## Syntax

 _expression_ . **SetFirstPriority**

 _expression_ A variable that represents a **Databar** object.


## Remarks

When you have multiple conditional formatting rules in a worksheet, if the rule was not previously set to priority "1", this method will cause the priority of all other existing rules on the worksheet to be increased by one.


 **Note**  Priority levels for conditional formatting rules are applied on a worksheet-level basis.


## See also


#### Concepts


[Databar Object](databar-object-excel.md)

