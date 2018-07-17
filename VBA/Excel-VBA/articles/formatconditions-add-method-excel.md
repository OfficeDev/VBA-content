---
title: FormatConditions.Add Method (Excel)
keywords: vbaxl10.chm510075
f1_keywords:
- vbaxl10.chm510075
ms.prod: excel
api_name:
- Excel.FormatConditions.Add
ms.assetid: 705f9ad4-2500-6607-19c0-6abd3f214d3e
ms.date: 06/08/2017
---


# FormatConditions.Add Method (Excel)

Adds a new conditional format.


## Syntax

 _expression_ . **Add**( **_Type_** , **_Operator_** , **_Formula1_** , **_Formula2_** )

 _expression_ A variable that represents a **FormatConditions** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[XlFormatConditionType](xlformatconditiontype-enumeration-excel.md)**|. Specifies whether the conditional format is based on a cell value or an expression.|
| _Operator_|Optional| **Variant**|The conditional format operator. Can be one of the following  **XlFormatConditionOperator** constants: **xlBetween** , **xlEqual** , **xlGreater** , **xlGreaterEqual** , **xlLess** , **xlLessEqual** , **xlNotBetween** , or **xlNotEqual** . If _Type_ is **xlExpression** , the _Operator_ argument is ignored.|
| _Formula1_|Optional| **Variant**|The value or expression associated with the conditional format. Can be a constant value, a string value, a cell reference, or a formula.|
| _Formula2_|Optional| **Variant**|The value or expression associated with the second part of the conditional format when  _Operator_ is **xlBetween** or **xlNotBetween** (otherwise, this argument is ignored). Can be a constant value, a string value, a cell reference, or a formula.|

### Return Value

A  **[FormatCondition](formatcondition-object-excel.md)** object that represents the new conditional format.


## Remarks

Use the  **[Modify](formatcondition-modify-method-excel.md)** method to modify an existing conditional format, or use the **[Delete](formatcondition-delete-method-excel.md)** method to delete an existing format before adding a new one.


## Example

This example adds a conditional format to cells E1:E10.


```vb
With Worksheets(1).Range("e1:e10").FormatConditions _ 
 .Add(xlCellValue, xlGreater, "=$a$1") 
 With .Borders 
 .LineStyle = xlContinuous 
 .Weight = xlThin 
 .ColorIndex = 6 
 End With 
 With .Font 
 .Bold = True 
 .ColorIndex = 3 
 End With 
End With
```


## See also


#### Concepts


[FormatConditions Object](formatconditions-object-excel.md)

