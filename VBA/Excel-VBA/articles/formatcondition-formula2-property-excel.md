---
title: FormatCondition.Formula2 Property (Excel)
keywords: vbaxl10.chm512077
f1_keywords:
- vbaxl10.chm512077
ms.prod: excel
api_name:
- Excel.FormatCondition.Formula2
ms.assetid: 2909d42d-7665-3406-8732-4a51034474c3
ms.date: 06/08/2017
---


# FormatCondition.Formula2 Property (Excel)

Returns the value or expression associated with the second part of a conditional format or data validation. Used only when the data validation conditional format  **[Operator](formatcondition-operator-property-excel.md)** property is **xlBetween** or **xlNotBetween** . Can be a constant value, a string value, a cell reference, or a formula. Read-only **String** .


## Syntax

 _expression_ . **Formula2**

 _expression_ A variable that represents a **FormatCondition** object.


## Example

This example changes the formula for conditional format one for cells E1:E10 if the formula specifies "between 5 and 10"


```vb
With Worksheets(1).Range("e1:e10").FormatConditions(1) 
 If .Operator = xlBetween And _ 
 .Formula1 = "5" And _ 
 .Formula2 = "10" Then 
 .Modify xlCellValue, xlLess, "10" 
 End If 
End With
```


## See also


#### Concepts


[FormatCondition Object](formatcondition-object-excel.md)

