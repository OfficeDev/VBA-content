---
title: IconCriterion.Type Property (Excel)
keywords: vbaxl10.chm814074
f1_keywords:
- vbaxl10.chm814074
ms.prod: excel
api_name:
- Excel.IconCriterion.Type
ms.assetid: bbe75bbb-42d1-7b71-7a7a-7c51e8c47cbc
ms.date: 06/08/2017
---


# IconCriterion.Type Property (Excel)

Returns one of the constants of the  **[XlConditionValueTypes](xlconditionvaluetypes-enumeration-excel.md)** enumeration, which specifies how the threshold value for an icon set is determined. Read-only.


## Syntax

 _expression_ . **Type**

 _expression_ A variable that represents an **IconCriterion** object.


## Remarks

The type of threshold value for an icon set can be a number, percent, formula, or percentile. Setting the type to percentile will use the Percentile function in Excel to determine the threshold value.


## See also


#### Concepts


[IconCriterion Object](iconcriterion-object-excel.md)

