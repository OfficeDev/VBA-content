---
title: FormatCondition.Priority Property (Excel)
keywords: vbaxl10.chm512087
f1_keywords:
- vbaxl10.chm512087
ms.prod: excel
api_name:
- Excel.FormatCondition.Priority
ms.assetid: 27d0a82a-b69b-de94-ff90-dbd3bd5a02fa
ms.date: 06/08/2017
---


# FormatCondition.Priority Property (Excel)

Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet.


## Syntax

 _expression_ . **Priority**

 _expression_ A variable that represents a **FormatCondition** object.


## Remarks

When setting the priority, the value must be a positive integer between 1 and the total number of conditional formatting rules on the worksheet. The priority must be a unique value for all rules on the worksheet, so changing the priority for the specified conditional formatting rule may cause the priority value of the other rules on the worksheet to be shifted.


## See also


#### Concepts


[FormatCondition Object](formatcondition-object-excel.md)

