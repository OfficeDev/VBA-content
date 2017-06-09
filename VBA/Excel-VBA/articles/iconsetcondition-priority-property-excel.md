---
title: IconSetCondition.Priority Property (Excel)
keywords: vbaxl10.chm812073
f1_keywords:
- vbaxl10.chm812073
ms.prod: excel
api_name:
- Excel.IconSetCondition.Priority
ms.assetid: c2f72c35-702a-cae7-ffde-ad7075c8dc75
ms.date: 06/08/2017
---


# IconSetCondition.Priority Property (Excel)

Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet.


## Syntax

 _expression_ . **Priority**

 _expression_ A variable that represents an **IconSetCondition** object.


## Remarks

When setting the priority, the value must be a positive integer between 1 and the total number of conditional formatting rules on the worksheet. The priority must be a unique value for all rules on the worksheet, so changing the priority for the specified conditional formatting rule may cause the priority value of the other rules on the worksheet to be shifted.


## See also


#### Concepts


[IconSetCondition Object](iconsetcondition-object-excel.md)

