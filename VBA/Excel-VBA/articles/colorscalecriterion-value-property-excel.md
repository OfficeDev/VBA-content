---
title: ColorScaleCriterion.Value Property (Excel)
keywords: vbaxl10.chm808075
f1_keywords:
- vbaxl10.chm808075
ms.prod: excel
api_name:
- Excel.ColorScaleCriterion.Value
ms.assetid: 829e876f-ca11-855d-bda5-a1c7f86eeb0f
ms.date: 06/08/2017
---


# ColorScaleCriterion.Value Property (Excel)

Returns or sets the minimum, midpoint, or maximum threshold value for a color scale conditional format. Read/write  **Variant** .


## Syntax

 _expression_ . **Value**

 _expression_ A variable that represents a **ColorScaleCriterion** object.


## Remarks

You can set the value only if the  **[ColorScaleCriterion.Type](colorscalecriterion-type-property-excel.md)** property for the conditional format is set to one of the following constants: **xlConditionValueNumber** , **xlConditionValuePercent** , **xlConditionValuePercentile** , or **xlConditionValueFormula** .

If the type of threshold is a formula, you can set the formula as a  **String** . The formula must return a single number.


## See also


#### Concepts


[ColorScaleCriterion Object](colorscalecriterion-object-excel.md)

