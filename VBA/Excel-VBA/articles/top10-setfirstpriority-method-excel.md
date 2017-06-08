---
title: Top10.SetFirstPriority Method (Excel)
keywords: vbaxl10.chm822084
f1_keywords:
- vbaxl10.chm822084
ms.prod: excel
api_name:
- Excel.Top10.SetFirstPriority
ms.assetid: 3523bdae-87ab-54f5-e6ff-a684592b88b7
ms.date: 06/08/2017
---


# Top10.SetFirstPriority Method (Excel)

Sets the priority value for this conditional formatting rule to "1" so that it will be evaluated before all other rules on the worksheet.


## Syntax

 _expression_ . **SetFirstPriority**

 _expression_ A variable that represents a **Top10** object.


## Remarks

When you have multiple conditional formatting rules in a worksheet, if the rule was not previously set to priority "1", this method will cause the priority of all other existing rules on the worksheet to be increased by one.


 **Note**  Priority levels for conditional formatting rules are applied on a worksheet-level basis.


## See also


#### Concepts


[Top10 Object](top10-object-excel.md)

