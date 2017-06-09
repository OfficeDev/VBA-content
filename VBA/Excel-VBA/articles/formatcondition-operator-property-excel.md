---
title: FormatCondition.Operator Property (Excel)
keywords: vbaxl10.chm512075
f1_keywords:
- vbaxl10.chm512075
ms.prod: excel
api_name:
- Excel.FormatCondition.Operator
ms.assetid: 943fd9c1-30b2-d2aa-e9fe-f243af6b1292
ms.date: 06/08/2017
---


# FormatCondition.Operator Property (Excel)

Returns a  **Long** value that represents the operator for the conditional format.


## Syntax

 _expression_ . **Operator**

 _expression_ A variable that represents a **FormatCondition** object.


## Example

This example changes the formula for conditional format one, for cells E1:E10 if the formula specifies "less than 5."


```vb
With Worksheets(1).Range("e1:e10").FormatConditions(1) 
 If .Operator = xlLess And .Formula1 = "5" Then 
 .Modify xlCellValue, xlBetween, "5", "15" 
 End If 
End With
```


## See also


#### Concepts


[FormatCondition Object](formatcondition-object-excel.md)

