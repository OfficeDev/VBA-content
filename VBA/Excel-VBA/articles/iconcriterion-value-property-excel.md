---
title: IconCriterion.Value Property (Excel)
keywords: vbaxl10.chm814075
f1_keywords:
- vbaxl10.chm814075
ms.prod: excel
api_name:
- Excel.IconCriterion.Value
ms.assetid: 5cb72b0b-1df2-dd47-932f-1454fda9f804
ms.date: 06/08/2017
---


# IconCriterion.Value Property (Excel)

Returns or sets the threshold value for an icon in a conditional format. Read/write  **Variant** .


## Syntax

 _expression_ . **Value**

 _expression_ A variable that represents an **IconCriterion** object.


## Remarks

You can set the value only if the  **[IconCriterion.Type](iconcriterion-type-property-excel.md)** property for the conditional format is set to one of the following constants: **xlConditionValueNumber** , **xlConditionValuePercent** , **xlConditionValuePercentile** , or **xlConditionValueFormula** .

If the type of threshold is a formula, you can set the formula as a  **String** . The formula must return a single number.


## See also


#### Concepts


[IconCriterion Object](iconcriterion-object-excel.md)

