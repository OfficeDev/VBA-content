---
title: FormatCondition.Formula1 Property (Excel)
keywords: vbaxl10.chm512076
f1_keywords:
- vbaxl10.chm512076
ms.prod: excel
api_name:
- Excel.FormatCondition.Formula1
ms.assetid: f711069a-0d4b-d70c-ed48-9c375ce29173
ms.date: 06/08/2017
---


# FormatCondition.Formula1 Property (Excel)

Returns the value or expression associated with the conditional format or data validation. Can be a constant value, a string value, a cell reference, or a formula. Read-only  **String** .


## Syntax

 _expression_ . **Formula1**

 _expression_ A variable that represents a **FormatCondition** object.


## Example

This example changes the formula for conditional format one for cells E1:E10 if the formula specifies "less than 5."


```vb
With Worksheets(1).Range("e1:e10").FormatConditions(1) 
 If .Operator = xlLess And .Formula1 = "5" Then 
 .Modify xlCellValue, xlLess, "10" 
 End If 
End With
```


## See also


#### Concepts


[FormatCondition Object](formatcondition-object-excel.md)

