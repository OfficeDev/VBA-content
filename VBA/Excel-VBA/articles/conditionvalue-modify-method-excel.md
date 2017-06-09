---
title: ConditionValue.Modify Method (Excel)
keywords: vbaxl10.chm804073
f1_keywords:
- vbaxl10.chm804073
ms.prod: excel
api_name:
- Excel.ConditionValue.Modify
ms.assetid: 3da6d850-7b7b-2419-b211-b18081c31e77
ms.date: 06/08/2017
---


# ConditionValue.Modify Method (Excel)

Modifies how the longest bar or shortest bar is evaluated for a data bar conditional formatting rule. 


## Syntax

 _expression_ . **Modify**( **_newtype_** , **_newvalue_** )

 _expression_ A variable that represents a **ConditionValue** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _newtype_|Required| **[XlConditionValueTypes](xlconditionvaluetypes-enumeration-excel.md)**|Specifies how the shortest bar or longest bar is evaluated. The default value is  **xlConditionLowestValue** for the shortest bar and **xlConditionHighestValue** for the longest bar.|
| _newvalue_|Optional| **Variant**|The value assigned to the shortest or longest data bar. Depending on the  _newtype_ argument, this can be a number or a formula that evaluates to a number.|

## Remarks

The following table describes the acceptable threshold values for each type of evaluation.



|**_newtype_ argument**|**_newvalue_ argument**|
|:-----|:-----|
|xlConditionLowestValue|argument is ignored|
|xlConditionHighestValue|argument is ignored|
|xlConditionValueNumber|any number|
|xlConditionValuePercent|any number between 0 and 100 |
|xlConditionValuePercentile|any number between 0 and 100|
|xlConditionValueFormula|a formula that returns a single number|

## See also


#### Concepts


[ConditionValue Object](conditionvalue-object-excel.md)

