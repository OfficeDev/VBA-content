---
title: AboveAverage.SetLastPriority Method (Excel)
keywords: vbaxl10.chm824083
f1_keywords:
- vbaxl10.chm824083
ms.prod: excel
api_name:
- Excel.AboveAverage.SetLastPriority
ms.assetid: e28605d2-338b-4efb-e7f0-f250bca85050
ms.date: 06/08/2017
---


# AboveAverage.SetLastPriority Method (Excel)

Sets the evaluation order for this conditional formatting rule so it is evaluated after all other rules on the worksheet.


## Syntax

 _expression_ . **SetLastPriority**

 _expression_ A variable that represents an **AboveAverage** object.


## Remarks

The actual value of the priority will be equal to the total number of conditional formatting rules on the worksheet. When you have multiple conditional formatting rules in a worksheet, this method will cause the priority of rules that had a priority value greater than this rule to be decreased by one.


 **Note**  Priority levels for conditional formatting rules are applied on a worksheet-level basis.


## See also


#### Concepts


[AboveAverage Object](aboveaverage-object-excel.md)

