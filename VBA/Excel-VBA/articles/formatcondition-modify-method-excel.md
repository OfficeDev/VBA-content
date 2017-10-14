---
title: FormatCondition.Modify Method (Excel)
keywords: vbaxl10.chm512082
f1_keywords:
- vbaxl10.chm512082
ms.prod: excel
api_name:
- Excel.FormatCondition.Modify
ms.assetid: a0dec05c-898d-87c9-9413-9182d31f6ed0
ms.date: 06/08/2017
---


# FormatCondition.Modify Method (Excel)

Modifies an existing conditional format.


## Syntax

 _expression_ . **Modify**( **_Type_** , **_Operator_** , **_Formula1_** , **_Formula2_** )

 _expression_ A variable that represents a **FormatCondition** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[XlFormatConditionType](xlformatconditiontype-enumeration-excel.md)**|Specifies whether the conditional format is based on a cell value or an expression.|
| _Operator_|Optional| **Variant**|An  **[XlFormatConditionOperator](xlformatconditionoperator-enumeration-excel.md)** value that represents the conditional format operator. This parameter is ignored if _Type_ is set to **xlExpression** .|
| _Formula1_|Optional| **Variant**|The value or expression associated with the conditional format. Can be a constant value, a string value, a cell reference, or a formula.|
| _Formula2_|Optional| **Variant**|The value or expression associated with the conditional format. Can be a constant value, a string value, a cell reference, or a formula..|

## Example

This example modifies an existing conditional format for cells E1:E10.


```vb
Worksheets(1).Range("e1:e10").FormatConditions(1) _ 
 .Modify xlCellValue, xlLess, "=$a$1"
```


## See also


#### Concepts


[FormatCondition Object](formatcondition-object-excel.md)

